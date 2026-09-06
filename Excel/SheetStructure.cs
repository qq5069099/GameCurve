using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using GameCurve.Models;

namespace GameCurve.Excel;

/// <summary>待应用到工作簿的结构改动类型。</summary>
public enum StructuralKind
{
    /// <summary>在指定物理行号处插入一行。</summary>
    InsertRow,
    /// <summary>删除指定物理行号。</summary>
    DeleteRow,
    /// <summary>在指定列索引处插入一列。</summary>
    InsertColumn,
    /// <summary>删除指定列索引。</summary>
    DeleteColumn
}

/// <summary>一条结构改动：作用于某工作表，Index 为物理行号或列索引。</summary>
public readonly record struct StructuralOp(string Sheet, StructuralKind Kind, int Index);

/// <summary>
/// 工作表行列结构的“平移式”编辑。所有改动都通过移动已有 Row/Cell 元素完成，
/// 因此原单元格样式、公式、共享字符串引用与宏都保持不变；只新增/删除必要的空行/空列。
/// </summary>
public static class SheetStructure
{
    /// <summary>在物理行号 rowNumber 处插入一个空行，其后的所有行下移一行。</summary>
    public static void InsertRow(WorksheetPart wsPart, int rowNumber)
    {
        var sheetData = GetSheetData(wsPart);
        var rows = ReadRows(sheetData);
        var target = new Row { RowIndex = (uint)Math.Max(1, rowNumber) };
        var result = new SortedDictionary<int, Row>();
        foreach (var (oldIndex, row) in rows)
        {
            int newIndex = oldIndex >= rowNumber ? oldIndex + 1 : oldIndex;
            ShiftRowNumber(row, oldIndex, newIndex);
            result[newIndex] = row;
        }
        result[Math.Max(1, rowNumber)] = target;
        WriteRows(sheetData, result);
    }

    /// <summary>删除物理行号 rowNumber 的行，其后的所有行上移一行。</summary>
    public static void DeleteRow(WorksheetPart wsPart, int rowNumber)
    {
        var sheetData = GetSheetData(wsPart);
        var rows = ReadRows(sheetData);
        var result = new SortedDictionary<int, Row>();
        foreach (var (oldIndex, row) in rows)
        {
            if (oldIndex == rowNumber) continue;
            int newIndex = oldIndex > rowNumber ? oldIndex - 1 : oldIndex;
            ShiftRowNumber(row, oldIndex, newIndex);
            result[newIndex] = row;
        }
        WriteRows(sheetData, result);
    }

    /// <summary>在列索引 colIndex 处插入一个空列，列号大于等于 colIndex 的单元格右移一列。</summary>
    public static void InsertColumn(WorksheetPart wsPart, int colIndex)
    {
        var sheetData = GetSheetData(wsPart);
        foreach (var row in sheetData.Elements<Row>())
            foreach (var cell in row.Elements<Cell>().ToList())
                ShiftCellColumn(cell, colIndex, +1);
        SortCellsPerRow(sheetData);
    }

    /// <summary>删除列索引 colIndex 的列，列号大于 colIndex 的单元格左移一列。</summary>
    public static void DeleteColumn(WorksheetPart wsPart, int colIndex)
    {
        var sheetData = GetSheetData(wsPart);
        foreach (var row in sheetData.Elements<Row>())
        {
            foreach (var cell in row.Elements<Cell>().ToList())
            {
                int col = GetColumn(cell);
                if (col == colIndex) { cell.Remove(); continue; }
                if (col > colIndex) ShiftCellColumn(cell, colIndex, -1);
            }
        }
        SortCellsPerRow(sheetData);
    }

    private static SheetData GetSheetData(WorksheetPart wsPart)
    {
        var sheetData = wsPart.Worksheet.GetFirstChild<SheetData>();
        if (sheetData == null)
        {
            sheetData = new SheetData();
            wsPart.Worksheet.AppendChild(sheetData);
        }
        return sheetData;
    }

    private static SortedDictionary<int, Row> ReadRows(SheetData sheetData)
    {
        var map = new SortedDictionary<int, Row>();
        foreach (var row in sheetData.Elements<Row>())
        {
            int index = (int)(row.RowIndex?.Value ?? 0);
            if (index <= 0) continue;
            map[index] = row;
        }
        return map;
    }

    private static void WriteRows(SheetData sheetData, SortedDictionary<int, Row> rows)
    {
        foreach (var row in sheetData.Elements<Row>().ToList()) row.Remove();
        foreach (var row in rows.Values)
            sheetData.AppendChild(row);
    }

    private static void ShiftRowNumber(Row row, int oldIndex, int newIndex)
    {
        row.RowIndex = (uint)Math.Max(1, newIndex);
        foreach (var cell in row.Elements<Cell>())
        {
            if (cell.CellReference?.Value == null) continue;
            int oldRow = CellHelper.RowNumberFromRef(cell.CellReference.Value);
            int col = CellHelper.LetterToColumnIndex(new string(cell.CellReference.Value.TakeWhile(char.IsLetter).ToArray()));
            cell.CellReference = CellHelper.ToCellReference(col, Math.Max(1, oldRow + (newIndex - oldIndex)));
        }
    }

    private static void ShiftCellColumn(Cell cell, int threshold, int delta)
    {
        if (cell.CellReference?.Value == null) return;
        int col = GetColumn(cell);
        if (col < threshold) return;
        int row = CellHelper.RowNumberFromRef(cell.CellReference.Value);
        cell.CellReference = CellHelper.ToCellReference(col + delta, row);
    }

    private static void SortCellsPerRow(SheetData sheetData)
    {
        foreach (var row in sheetData.Elements<Row>())
        {
            var cells = row.Elements<Cell>()
                .OrderBy(GetColumn)
                .ToList();
            foreach (var cell in row.Elements<Cell>().ToList()) cell.Remove();
            foreach (var cell in cells) row.AppendChild(cell);
        }
    }

    private static int GetColumn(Cell cell)
        => cell.CellReference?.Value == null
            ? -1
            : CellHelper.LetterToColumnIndex(new string(cell.CellReference.Value.TakeWhile(char.IsLetter).ToArray()));
}
