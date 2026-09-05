using System.Globalization;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace GameCurve.Excel;

/// <summary>
/// OpenXml 单元格读写辅助。所有操作都只改目标单元格，保留其余内容、格式与宏。
/// </summary>
public static class CellHelper
{
    private static readonly Dictionary<uint, string> SharedStringCache = new();

    public static void SetSharedStrings(SharedStringTable? table)
    {
        SharedStringCache.Clear();
        if (table == null) return;
        int i = 0;
        foreach (var item in table.Elements<SharedStringItem>())
        {
            SharedStringCache[(uint)i++] = item.InnerText;
        }
    }

    /// <summary>把单元格解析成字符串文本（共享字符串/内联字符串/布尔/数值统一转文本）。</summary>
    public static string? GetCellValue(Cell cell)
    {
        if (cell == null) return null;
        var t = cell.DataType?.Value;
        var v = cell.CellValue?.Text;
        if (t == CellValues.SharedString)
        {
            if (v != null && uint.TryParse(v, out var idx) && SharedStringCache.TryGetValue(idx, out var s))
                return s;
            return v;
        }
        if (t == CellValues.InlineString)
            return cell.InlineString?.InnerText;
        if (t == CellValues.Boolean)
            return v == "1" ? "true" : (v == "0" ? "false" : v);
        return v;
    }

    /// <summary>尝试把某列原始文本解析为 double，跳过空值与不可解析项。</summary>
    public static bool TryParseDouble(string? text, out double value)
    {
        value = 0;
        if (string.IsNullOrWhiteSpace(text)) return false;
        var s = text.Trim();
        if (double.TryParse(s, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out value))
            return true;
        // 个别表可能用逗号做小数，尝试本地化
        if (double.TryParse(s, NumberStyles.Float, CultureInfo.CurrentCulture, out value))
            return true;
        return false;
    }

    public static string ToCellReference(int columnIndex, int rowNumber) => $"{ColumnIndexToLetter(columnIndex)}{rowNumber}";

    public static string ColumnIndexToLetter(int columnIndex)
    {
        if (columnIndex < 0) return "";
        var sb = new StringBuilder();
        int c = columnIndex;
        while (c >= 0)
        {
            sb.Insert(0, (char)('A' + c % 26));
            c = c / 26 - 1;
        }
        return sb.ToString();
    }

    public static int LetterToColumnIndex(string letters)
    {
        int idx = 0;
        foreach (var ch in letters.ToUpperInvariant())
        {
            idx = idx * 26 + (ch - 'A' + 1);
        }
        return idx - 1;
    }

    public static int RowNumberFromRef(string cellRef)
    {
        int i = 0;
        while (i < cellRef.Length && char.IsLetter(cellRef[i])) i++;
        return int.TryParse(cellRef[i..], out var r) ? r : 0;
    }

    public static int GetColumnIndex(Cell cell)
    {
        var refText = cell.CellReference?.Value ?? "";
        return LetterToColumnIndex(new string(refText.TakeWhile(char.IsLetter).ToArray()));
    }

    /// <summary>按引用获取单元格；不存在则创建并插入到正确排序位置。</summary>
    public static Cell GetOrCreateCell(WorksheetPart wsPart, string cellRef)
    {
        var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            wsPart.Worksheet.AppendChild(sheetData);
        }

        // 先查找
        var existing = sheetData.Descendants<Cell>().FirstOrDefault(c => string.Equals(c.CellReference?.Value, cellRef, StringComparison.OrdinalIgnoreCase));
        if (existing != null) return existing;

        int rowNumber = RowNumberFromRef(cellRef);
        int colIndex = LetterToColumnIndex(new string(cellRef.TakeWhile(char.IsLetter).ToArray()));

        var row = GetOrCreateRow(sheetData, rowNumber);

        var cell = new Cell
        {
            CellReference = cellRef,
            DataType = CellValues.Number,
            CellValue = new CellValue("0")
        };

        // 按列序插入
        var cells = row.Elements<Cell>().ToList();
        int insertAt = cells.Count;
        for (int i = 0; i < cells.Count; i++)
        {
            var r = cells[i].CellReference?.Value;
            if (r == null) continue;
            int c = LetterToColumnIndex(new string(r.TakeWhile(char.IsLetter).ToArray()));
            if (c > colIndex)
            {
                insertAt = i;
                break;
            }
        }
        row.InsertAt(cell, insertAt);
        return cell;
    }

    private static Row GetOrCreateRow(SheetData sheetData, int rowNumber)
    {
        var row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex?.Value == (uint)rowNumber);
        if (row != null) return row;

        row = new Row { RowIndex = (uint)rowNumber };
        int insertAt = sheetData.Elements<Row>().Count();
        int i = 0;
        foreach (var r in sheetData.Elements<Row>())
        {
            var rr = r.RowIndex?.Value ?? 0;
            if (rr > rowNumber)
            {
                insertAt = i;
                break;
            }
            i++;
        }
        sheetData.InsertAt(row, insertAt);
        return row;
    }

    /// <summary>设置单元格为数值。integerMode=true 时四舍五入为整数，否则保留指定小数位。</summary>
    public static void SetNumericValue(WorksheetPart wsPart, string cellRef, double value, bool integerMode, int decimalPlaces)
    {
        string text;
        if (integerMode)
        {
            long iv = (long)Math.Round(value);
            text = iv.ToString(CultureInfo.InvariantCulture);
        }
        else
        {
            text = value.ToString("F" + Math.Max(0, decimalPlaces), CultureInfo.InvariantCulture).TrimEnd('0').TrimEnd('.');
            if (text == "-0" || text == "") text = "0";
        }

        var cell = GetOrCreateCell(wsPart, cellRef);

        // 清除旧的内联字符串缓存，避免混入
        var inline = cell.InlineString;
        if (inline != null) inline.Remove();

        cell.CellValue = new CellValue(text);
        cell.DataType = CellValues.Number;
    }

    /// <summary>把单元格写为文本（内联字符串，保留原单元格样式）。</summary>
    public static void SetStringValue(WorksheetPart wsPart, string cellRef, string text)
    {
        var cell = GetOrCreateCell(wsPart, cellRef);
        cell.InlineString?.Remove();
        cell.CellValue?.Remove();
        cell.CellFormula?.Remove();
        cell.DataType = CellValues.InlineString;
        cell.InlineString = new InlineString(new Text(text));
    }
}
