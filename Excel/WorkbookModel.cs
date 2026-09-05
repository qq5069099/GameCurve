using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using GameCurve.Models;

namespace GameCurve.Excel;

/// <summary>一条待写回单元格的指令。</summary>
public readonly record struct CellWrite(string SheetName, string CellRef, double Value, bool IntegerMode, int DecimalPlaces);
/// <summary>写文本（非数值）单元格。</summary>
public readonly record struct CellWriteString(string SheetName, string CellRef, string Value);

/// <summary>
/// 工作簿入口：负责枚举工作表、加载某表快照、识别数值列、批量写回单元格。
/// 采用“读时只读打开、写时编辑打开并保存”的方式，避免长期占用文件导致 Excel 无法保存。
/// </summary>
public sealed class WorkbookModel : IDisposable
{
    public string Path { get; private set; } = "";
    public int HeaderRow { get; set; } = 1;

    /// <summary>文件最后写入时间（UTC），用于检测外部修改。</summary>
    public DateTime LastModified { get; private set; }

    private readonly Dictionary<string, SheetSnapshot> _cache = new(StringComparer.OrdinalIgnoreCase);

    public IReadOnlyList<string> SheetNames { get; private set; } = Array.Empty<string>();

    public void Open(string path)
    {
        Path = path;
        _cache.Clear();
        using var doc = SpreadsheetDocument.Open(path, false);
        var wb = doc.WorkbookPart!.Workbook;
        SheetNames = (wb.Sheets?.Elements<Sheet>().Select(s => s.Name!.Value!).ToList() ?? new List<string>());
        LastModified = File.GetLastWriteTimeUtc(path);
    }

    /// <summary>重新读取文件元信息（用于外部变更后重载）。</summary>
    public void RefreshMeta()
    {
        LastModified = File.GetLastWriteTimeUtc(Path);
        _cache.Clear();
    }

    /// <summary>加载某工作表并生成列元信息与网格快照。</summary>
    public SheetSnapshot LoadSheet(string sheetName)
    {
        if (_cache.TryGetValue(sheetName, out var cached)) return cached;

        using var doc = SpreadsheetDocument.Open(Path, false);
        var wbPart = doc.WorkbookPart!;
        CellHelper.SetSharedStrings(wbPart.SharedStringTablePart?.SharedStringTable);
        var wsPart = GetWorksheetPart(wbPart, sheetName);
        var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();

        // 读取单元格样式表，用于还原原始水平对齐
        var cellFormats = wbPart.WorkbookStylesPart?.Stylesheet?.CellFormats?.Elements<CellFormat>().ToList();
        static HorizontalAlign? AlignOf(Cell cell, List<CellFormat>? formats)
        {
            if (formats == null) return null;
            uint si = cell.StyleIndex?.Value ?? 0;
            if (si >= formats.Count) return null;
            var hv = formats[(int)si].Alignment?.Horizontal?.Value;
            if (hv == HorizontalAlignmentValues.Center || hv == HorizontalAlignmentValues.CenterContinuous)
                return HorizontalAlign.Center;
            if (hv == HorizontalAlignmentValues.Left) return HorizontalAlign.Left;
            if (hv == HorizontalAlignmentValues.Right) return HorizontalAlign.Right;
            return null;
        }

        // 收集所有行
        var rows = new SortedDictionary<int, List<Cell>>();
        int maxRow = 0, maxCol = 0;
        foreach (var row in sheetData!.Elements<Row>())
        {
            uint r = row.RowIndex?.Value ?? 0;
            if (r == 0) continue;
            int rowNum = (int)r;
            var list = new List<Cell>();
            foreach (var cell in row.Elements<Cell>())
            {
                if (cell.CellReference?.Value == null) continue;
                int colIdx = CellHelper.GetColumnIndex(cell);
                list.Add(cell);
                if (colIdx + 1 > maxCol) maxCol = colIdx + 1;
            }
            rows[rowNum] = list;
            if (rowNum > maxRow) maxRow = rowNum;
        }

        int headerRow = DetectHeaderRow(rows);
        string?[] headerRaw = new string?[maxCol];
        if (rows.TryGetValue(headerRow, out var hc))
        {
            foreach (var cell in hc)
            {
                int colIdx = CellHelper.GetColumnIndex(cell);
                if (colIdx >= 0 && colIdx < maxCol)
                    headerRaw[colIdx] = CellHelper.GetCellValue(cell);
            }
        }

        var snap = new SheetSnapshot
        {
            SheetName = sheetName,
            HeaderRow = headerRow,
            MaxRow = maxRow,
            ColumnCount = maxCol
        };

        // 建立列元信息
        var samples = PrecomputeSamples(rows, maxRow, maxCol, headerRow);
        for (int c = 0; c < maxCol; c++)
        {
            var raw = headerRaw[c];
            if (raw == null) raw = "";
            var name = HeaderParser.ParseName(raw);
            var type = HeaderParser.ParseType(raw);
            var label = HeaderParser.ParseLabel(raw);
            var explicitNumeric = HeaderParser.IsExplicitNumericScalar(raw, type);
            var explicitInteger = HeaderParser.IsExplicitInteger(type);
            var (nonEmpty, numericCount, allInteger) = samples[c];
            bool isNumeric = explicitNumeric || (string.IsNullOrEmpty(type) && nonEmpty > 0 && (double)numericCount / nonEmpty >= 0.7);
            bool isInteger = explicitInteger || (string.IsNullOrEmpty(type) && allInteger && isNumeric);

            snap.Columns.Add(new ColumnMeta
            {
                ColumnIndex = c,
                Letter = CellHelper.ColumnIndexToLetter(c),
                Name = name,
                Label = label,
                Type = type,
                HeaderRaw = raw,
                IsEmpty = string.IsNullOrWhiteSpace(raw),
                IsNumericScalar = isNumeric,
                IsInteger = isInteger,
                NonEmptyCount = nonEmpty,
                NumericCount = numericCount,
                TotalRows = Math.Max(0, maxRow - headerRow)
            });
        }

        // 建立网格
        var alignVotes = new Dictionary<int, Dictionary<HorizontalAlign, int>>();
        for (int r = headerRow + 1; r <= maxRow; r++)
        {
            var rowCells = rows.TryGetValue(r, out var rc) ? rc : new List<Cell>();
            var line = new string?[maxCol];
            foreach (var cell in rowCells)
            {
                int colIdx = CellHelper.GetColumnIndex(cell);
                if (colIdx >= 0 && colIdx < maxCol)
                {
                    line[colIdx] = CellHelper.GetCellValue(cell);
                    var a = AlignOf(cell, cellFormats);
                    if (a != null)
                    {
                        if (!alignVotes.TryGetValue(colIdx, out var votes)) { votes = new(); alignVotes[colIdx] = votes; }
                        votes[a.Value] = votes.GetValueOrDefault(a.Value) + 1;
                    }
                }
            }
            snap.Grid.Add(line);
            snap.RowNumbers.Add(r);
        }

        // 每列取出现次数最多的对齐方式，作为整列默认显示
        for (int c = 0; c < maxCol; c++)
        {
            var align = HorizontalAlign.Default;
            if (alignVotes.TryGetValue(c, out var votes) && votes.Count > 0)
                align = votes.OrderByDescending(kv => kv.Value).First().Key;
            snap.ColumnAlignments.Add(align);
        }

        _cache[sheetName] = snap;
        HeaderRow = headerRow;
        return snap;
    }

    /// <summary>自动识别“表头行”：前 10 行里含字段定义标记（: &lt; 等）最多的行；否则默认第 1 行。</summary>
    private static int DetectHeaderRow(SortedDictionary<int, List<Cell>> rows)
    {
        int best = 1, bestScore = 0;
        foreach (var kv in rows)
        {
            int r = kv.Key;
            if (r > 10) break;
            int score = 0;
            foreach (var cell in kv.Value)
            {
                var text = CellHelper.GetCellValue(cell);
                if (string.IsNullOrWhiteSpace(text)) continue;
                var t = text.Trim();
                if (CellHelper.TryParseDouble(t, out _)) continue;
                if (t.StartsWith("{") || t.StartsWith("[")) continue; // json 数据行
                bool isDef = t.Contains(':') || t.Contains('<') || t.Contains(".json") ||
                             t.Contains(".all") || t.StartsWith('*') || t.StartsWith('?') || t.StartsWith('#');
                if (isDef) score++;
            }
            if (score > bestScore) { bestScore = score; best = r; }
        }
        return bestScore > 0 ? best : 1;
    }

    /// <summary>批量写回单元格，返回是否成功；失败时输出原因（如文件被占用）。</summary>
    public bool TryWriteCells(IReadOnlyList<CellWrite> writes, out string? error)
    {
        error = null;
        if (writes.Count == 0) return true;
        // 按工作表分组，避免重复打开
        var groups = writes.GroupBy(w => w.SheetName);
        try
        {
            using var doc = SpreadsheetDocument.Open(Path, true);
            var wbPart = doc.WorkbookPart!;
            foreach (var group in groups)
            {
                var wsPart = GetWorksheetPart(wbPart, group.Key);
                foreach (var w in group)
                    CellHelper.SetNumericValue(wsPart, w.CellRef, w.Value, w.IntegerMode, w.DecimalPlaces);
            }
            doc.Save();
            LastModified = File.GetLastWriteTimeUtc(Path);
            return true;
        }
        catch (IOException ex)
        {
            error = "文件被占用，无法保存。请关闭 Excel 后重试。" + ex.Message;
            return false;
        }
        catch (Exception ex)
        {
            error = "保存失败：" + ex.Message;
            return false;
        }
    }

    /// <summary>批量写回文本单元格（非数值），保留宏。</summary>
    public bool TryWriteCellsString(IReadOnlyList<CellWriteString> writes, out string? error)
    {
        error = null;
        if (writes.Count == 0) return true;
        var groups = writes.GroupBy(w => w.SheetName);
        try
        {
            using var doc = SpreadsheetDocument.Open(Path, true);
            var wbPart = doc.WorkbookPart!;
            foreach (var group in groups)
            {
                var wsPart = GetWorksheetPart(wbPart, group.Key);
                foreach (var w in group)
                    CellHelper.SetStringValue(wsPart, w.CellRef, w.Value);
            }
            doc.Save();
            LastModified = File.GetLastWriteTimeUtc(Path);
            return true;
        }
        catch (IOException ex)
        {
            error = "文件被占用，无法保存。请关闭 Excel 后重试。" + ex.Message;
            return false;
        }
        catch (Exception ex)
        {
            error = "保存失败：" + ex.Message;
            return false;
        }
    }

    /// <summary>保存到新路径（另存为）。</summary>
    public bool TrySaveAs(string destination, out string? error)
    {
        error = null;
        try
        {
            File.Copy(Path, destination, true);
            return true;
        }
        catch (Exception ex)
        {
            error = "另存为失败：" + ex.Message;
            return false;
        }
    }

    private static WorksheetPart GetWorksheetPart(WorkbookPart wbPart, string sheetName)
    {
        var sheet = wbPart.Workbook.Sheets!.Elements<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, sheetName, StringComparison.OrdinalIgnoreCase))
            ?? throw new InvalidOperationException($"找不到工作表 {sheetName}");
        return (WorksheetPart)wbPart.GetPartById(sheet.Id!);
    }

    /// <summary>预先统计每列的非空/数值/是否全整数（最多采样前 400 行）。</summary>
    private static (int NonEmpty, int Numeric, bool AllInteger)[] PrecomputeSamples(SortedDictionary<int, List<Cell>> rows, int maxRow, int maxCol, int headerRow)
    {
        var result = new (int, int, bool)[maxCol];
        for (int c = 0; c < maxCol; c++) result[c] = (0, 0, true);

        int count = 0;
        foreach (var kv in rows)
        {
            if (kv.Key <= headerRow) continue; // 跳过表头行
            if (count++ > 400) break;
            foreach (var cell in kv.Value)
            {
                if (cell.CellReference?.Value == null) continue;
                int colIdx = CellHelper.GetColumnIndex(cell);
                if (colIdx < 0 || colIdx >= maxCol) continue;
                var text = CellHelper.GetCellValue(cell);
                if (string.IsNullOrWhiteSpace(text)) continue;
                var (ne, nu, ai) = result[colIdx];
                ne++;
                if (CellHelper.TryParseDouble(text, out var d))
                {
                    nu++;
                    if (d != Math.Truncate(d)) ai = false;
                }
                result[colIdx] = (ne, nu, ai);
            }
        }
        return result;
    }

    public void Dispose()
    {
        _cache.Clear();
    }
}
