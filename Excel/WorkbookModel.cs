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

    /// <summary>用于持久化各工作表列显示顺序的非常隐藏内部配置表名。</summary>
    private const string OrderSheetName = "__GameCurve__";
    /// <summary>打开时从配置表读入的列顺序（按工作表名缓存）。</summary>
    private readonly Dictionary<string, int[]> _columnOrders = new(StringComparer.OrdinalIgnoreCase);

    public IReadOnlyList<string> SheetNames { get; private set; } = Array.Empty<string>();

    public void Open(string path)
    {
        Path = path;
        _cache.Clear();
        using var doc = SpreadsheetDocument.Open(path, false);
        var wbPart = doc.WorkbookPart!;
        var wb = wbPart.Workbook;
        // 内部配置表“__GameCurve__”不出现在用户可见的工作表列表中
        SheetNames = (wb.Sheets?.Elements<Sheet>()
            .Where(s => !string.Equals(s.Name?.Value, OrderSheetName, StringComparison.OrdinalIgnoreCase))
            .Select(s => s.Name!.Value!).ToList() ?? new List<string>());
        LoadColumnOrders(wbPart);
        LastModified = File.GetLastWriteTimeUtc(path);
    }

    /// <summary>重新读取文件元信息（用于外部变更后重载）。</summary>
    public void RefreshMeta()
    {
        LastModified = File.GetLastWriteTimeUtc(Path);
        _cache.Clear();
    }

    /// <summary>从内部配置表读入各工作表列顺序。</summary>
    private void LoadColumnOrders(WorkbookPart wbPart)
    {
        _columnOrders.Clear();
        var sheet = wbPart.Workbook.Sheets?.Elements<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, OrderSheetName, StringComparison.OrdinalIgnoreCase));
        if (sheet == null) return;
        var wsPart = (WorksheetPart)wbPart.GetPartById(sheet.Id!);
        var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData == null) return;
        CellHelper.SetSharedStrings(wbPart.SharedStringTablePart?.SharedStringTable);
        foreach (var row in sheetData.Elements<Row>())
        {
            var cells = row.Elements<Cell>().ToList();
            if (cells.Count < 2) continue;
            var name = CellHelper.GetCellValue(cells[0]);
            var text = CellHelper.GetCellValue(cells[1]);
            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(text)) continue;
            var order = ParseColumnOrder(text);
            if (order != null && order.Length > 0) _columnOrders[name] = order;
        }
    }

    private static int[]? ParseColumnOrder(string text)
    {
        var parts = text.Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);
        var list = new List<int>();
        foreach (var p in parts)
            if (int.TryParse(p, out var v) && v >= 0) list.Add(v);
        return list.Count > 0 ? list.Distinct().ToArray() : null;
    }

    /// <summary>返回某工作表的列显示顺序（物理列索引数组）；未记录时返回 null。</summary>
    public IReadOnlyList<int>? GetColumnOrder(string sheetName)
        => _columnOrders.TryGetValue(sheetName, out var o) ? o : null;

    /// <summary>把某工作表的列显示顺序写入内部配置表（非常隐藏），随文件持久化。</summary>
    public bool TryWriteColumnOrder(string sheetName, IReadOnlyList<int> order, out string? error)
    {
        error = null;
        if (order == null || order.Count == 0) return true;
        try
        {
            _columnOrders[sheetName] = order.Distinct().ToArray();
            using var doc = SpreadsheetDocument.Open(Path, true);
            var wbPart = doc.WorkbookPart!;
            var wsPart = GetOrCreateOrderSheet(wbPart);
            var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null)
            {
                sheetData = new SheetData();
                wsPart.Worksheet.AppendChild(sheetData);
            }
            // 每次全量重写，确保多个工作表的顺序都被保留且不重复
            sheetData.RemoveAllChildren();
            int rowNum = 1;
            foreach (var kv in _columnOrders)
            {
                var row = new Row { RowIndex = (uint)rowNum };
                row.Append(
                    MakeInlineCell("A" + rowNum, kv.Key),
                    MakeInlineCell("B" + rowNum, string.Join(",", kv.Value)));
                sheetData.AppendChild(row);
                rowNum++;
            }
            wsPart.Worksheet.Save();
            doc.Save();
            LastModified = File.GetLastWriteTimeUtc(Path);
            return true;
        }
        catch (IOException ex)
        {
            error = "文件被占用，无法保存列顺序。请关闭 Excel 后重试。" + ex.Message;
            return false;
        }
        catch (Exception ex)
        {
            error = "保存列顺序失败：" + ex.Message;
            return false;
        }
    }

    private static Cell MakeInlineCell(string cellRef, string text) => new()
    {
        CellReference = cellRef,
        DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text(text))
    };

    private static WorksheetPart GetOrCreateOrderSheet(WorkbookPart wbPart)
    {
        var sheet = wbPart.Workbook.Sheets?.Elements<Sheet>()
            .FirstOrDefault(s => string.Equals(s.Name?.Value, OrderSheetName, StringComparison.OrdinalIgnoreCase));
        if (sheet != null) return (WorksheetPart)wbPart.GetPartById(sheet.Id!);

        var wsPart = wbPart.AddNewPart<WorksheetPart>();
        if (wsPart.Worksheet == null) wsPart.Worksheet = new Worksheet();
        var sheets = wbPart.Workbook.GetFirstChild<Sheets>();
        uint sheetId = 1;
        if (sheets != null && sheets.Elements<Sheet>().Any())
            sheetId = (uint)(sheets.Elements<Sheet>().Max(s => s.SheetId?.Value ?? 0) + 1);
        sheets ??= wbPart.Workbook.AppendChild(new Sheets());
        sheets.AppendChild(new Sheet
        {
            Name = OrderSheetName,
            SheetId = sheetId,
            Id = wbPart.GetIdOfPart(wsPart),
            State = SheetStateValues.VeryHidden
        });
        return wsPart;
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
        var headerAligns = new HorizontalAlign[maxCol];
        Array.Fill(headerAligns, HorizontalAlign.Default);
        if (rows.TryGetValue(headerRow, out var hc))
        {
            foreach (var cell in hc)
            {
                int colIdx = CellHelper.GetColumnIndex(cell);
                if (colIdx >= 0 && colIdx < maxCol)
                {
                    headerRaw[colIdx] = CellHelper.GetCellValue(cell);
                    headerAligns[colIdx] = AlignOf(cell, cellFormats) ?? HorizontalAlign.Default;
                }
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
            snap.HeaderAlignments.Add(headerAligns[c]);
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

    /// <summary>
    /// 批量写回表头单元格的水平对齐样式。只为已存在的表头单元格新建/复用样式，
    /// 不会修改其它单元格或共享样式，因此不会影响数据列的外观与格式。
    /// </summary>
    public bool TryWriteHeaderAlignments(IReadOnlyList<(string Sheet, int Col, int HeaderRow, HorizontalAlign Align)> writes, out string? error)
    {
        error = null;
        if (writes.Count == 0) return true;
        try
        {
            using var doc = SpreadsheetDocument.Open(Path, true);
            var wbPart = doc.WorkbookPart!;
            var stylesPart = wbPart.WorkbookStylesPart;
            if (stylesPart?.Stylesheet == null)
                throw new InvalidOperationException("工作簿缺少样式表，无法写入表头对齐。");
            var stylesheet = stylesPart.Stylesheet;
            var cellFormats = stylesheet.GetFirstChild<CellFormats>();
            if (cellFormats == null)
                cellFormats = stylesheet.AppendChild(new CellFormats());

            var groups = writes.GroupBy(w => w.Sheet);
            foreach (var group in groups)
            {
                var wsPart = GetWorksheetPart(wbPart, group.Key);
                foreach (var w in group)
                {
                    string cellRef = CellHelper.ToCellReference(w.Col, w.HeaderRow);
                    var cell = CellHelper.FindCell(wsPart, cellRef);
                    if (cell == null) continue; // 表头单元格不存在时不强行新建，避免写入多余单元格

                    uint idx = EnsureAlignmentCellFormat(cellFormats, cell, w.Align);
                    cell.StyleIndex = idx;
                }
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

    /// <summary>
    /// 批量写回某列数据单元格的水平对齐样式。只处理已存在的数据单元格（表头行以下），
    /// 不会新建空单元格，也不会修改其它列或共享样式。
    /// </summary>
    public bool TryWriteColumnAlignments(IReadOnlyList<(string Sheet, int Col, int HeaderRow, int MaxRow, HorizontalAlign Align)> writes, out string? error)
    {
        error = null;
        if (writes.Count == 0) return true;
        try
        {
            using var doc = SpreadsheetDocument.Open(Path, true);
            var wbPart = doc.WorkbookPart!;
            var stylesPart = wbPart.WorkbookStylesPart;
            if (stylesPart?.Stylesheet == null)
                throw new InvalidOperationException("工作簿缺少样式表，无法写入列内容对齐。");
            var stylesheet = stylesPart.Stylesheet;
            var cellFormats = stylesheet.GetFirstChild<CellFormats>();
            if (cellFormats == null)
                cellFormats = stylesheet.AppendChild(new CellFormats());

            var groups = writes.GroupBy(w => w.Sheet);
            foreach (var group in groups)
            {
                var wsPart = GetWorksheetPart(wbPart, group.Key);
                int rangeStart = group.Min(w => w.HeaderRow) + 1;
                int rangeEnd = group.Max(w => w.MaxRow);
                var byCol = CellHelper.GetCellsByColumn(wsPart, rangeStart, rangeEnd);
                foreach (var w in group)
                {
                    if (!byCol.TryGetValue(w.Col, out var cells)) continue;
                    foreach (var cell in cells)
                    {
                        uint idx = EnsureAlignmentCellFormat(cellFormats, cell, w.Align);
                        cell.StyleIndex = idx;
                    }
                }
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

    /// <summary>
    /// 在样式表里找到与“当前单元格样式 + 目标水平对齐”一致的 CellFormat；没有则新建一个克隆，
    /// 只改水平对齐，保留字体/填充/边框/数字格式。返回该样式的索引。
    /// </summary>
    private static uint EnsureAlignmentCellFormat(CellFormats cellFormats, Cell cell, HorizontalAlign align)
    {
        var existing = cellFormats.Elements<CellFormat>().ToList();
        var targetHv = align switch
        {
            HorizontalAlign.Left => HorizontalAlignmentValues.Left,
            HorizontalAlign.Right => HorizontalAlignmentValues.Right,
            _ => HorizontalAlignmentValues.Center
        };

        uint baseIndex = cell.StyleIndex?.Value ?? 0;
        CellFormat? baseFmt = baseIndex < existing.Count ? existing[(int)baseIndex] : null;
        var desiredAlign = BuildAlignment(baseFmt, targetHv);

        // 优先复用完全一致的样式
        for (int i = 0; i < existing.Count; i++)
            if (AlignmentFormatMatches(existing[i], baseFmt, desiredAlign))
                return (uint)i;

        // 否则克隆基准样式，仅替换水平对齐
        var clone = baseFmt == null ? new CellFormat() : (CellFormat)baseFmt.CloneNode(true);
        clone.Alignment?.Remove();
        clone.Alignment = desiredAlign;
        cellFormats.AppendChild(clone);
        return (uint)(cellFormats.Count() - 1);
    }

    private static bool AlignmentFormatMatches(CellFormat fmt, CellFormat? baseFmt, Alignment desiredAlign)
    {
        if (baseFmt != null)
        {
            if (fmt.NumberFormatId?.Value != baseFmt.NumberFormatId?.Value) return false;
            if (fmt.FontId?.Value != baseFmt.FontId?.Value) return false;
            if (fmt.FillId?.Value != baseFmt.FillId?.Value) return false;
            if (fmt.BorderId?.Value != baseFmt.BorderId?.Value) return false;
            if (fmt.ApplyNumberFormat?.Value != baseFmt.ApplyNumberFormat?.Value) return false;
            if (fmt.ApplyFont?.Value != baseFmt.ApplyFont?.Value) return false;
            if (fmt.ApplyFill?.Value != baseFmt.ApplyFill?.Value) return false;
            if (fmt.ApplyBorder?.Value != baseFmt.ApplyBorder?.Value) return false;
        }
        return fmt.Alignment?.Horizontal == desiredAlign.Horizontal
            && fmt.Alignment?.Vertical == desiredAlign.Vertical;
    }

    private static Alignment BuildAlignment(CellFormat? baseFmt, HorizontalAlignmentValues hv)
    {
        var old = baseFmt?.Alignment;
        return new Alignment
        {
            Horizontal = hv,
            Vertical = old?.Vertical,
            WrapText = old?.WrapText,
            TextRotation = old?.TextRotation,
            Indent = old?.Indent,
            ReadingOrder = old?.ReadingOrder,
            ShrinkToFit = old?.ShrinkToFit
        };
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

    /// <summary>
    /// 按顺序批量应用结构改动（插入/删除行列），一次打开并保存，失败返回原因。
    /// 改动只平移现有单元格，不动样式与公式，宏文件（vbaProject.bin）原样保留。
    /// </summary>
    public bool TryApplyStructure(IReadOnlyList<StructuralOp> ops, out string? error)
    {
        error = null;
        if (ops.Count == 0) return true;
        try
        {
            using var doc = SpreadsheetDocument.Open(Path, true);
            var wbPart = doc.WorkbookPart!;
            foreach (var op in ops)
            {
                var wsPart = GetWorksheetPart(wbPart, op.Sheet);
                switch (op.Kind)
                {
                    case StructuralKind.InsertRow:
                        SheetStructure.InsertRow(wsPart, op.Index);
                        break;
                    case StructuralKind.DeleteRow:
                        SheetStructure.DeleteRow(wsPart, op.Index);
                        break;
                    case StructuralKind.InsertColumn:
                        SheetStructure.InsertColumn(wsPart, op.Index);
                        break;
                    case StructuralKind.DeleteColumn:
                        SheetStructure.DeleteColumn(wsPart, op.Index);
                        break;
                }
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
