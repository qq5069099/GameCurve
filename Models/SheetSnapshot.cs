using GameCurve.Excel;

namespace GameCurve.Models;

/// <summary>单元格水平对齐方式（来自 Excel 单元格样式）。</summary>
public enum HorizontalAlign
{
    Default = 0,
    Left = 1,
    Center = 2,
    Right = 3
}

/// <summary>
/// 某个工作表读入内存后的快照：所有单元格原始文本 + 列元信息 + 数据行号。
/// </summary>
public sealed class SheetSnapshot
{
    public string SheetName { get; set; } = "";
    public int HeaderRow { get; set; } = 1;

    /// <summary>最后一行的物理行号。</summary>
    public int MaxRow { get; set; }

    /// <summary>列数。</summary>
    public int ColumnCount { get; set; }

    public List<ColumnMeta> Columns { get; } = new();

    /// <summary>每行网格原始文本，索引 [i][col]，i 对应数据行顺序。</summary>
    public List<string?[]> Grid { get; } = new();

    /// <summary>每行对应的物理行号（HeaderRow+1 起）。</summary>
    public List<int> RowNumbers { get; } = new();

    /// <summary>每列的水平对齐方式（按列索引，长度同 ColumnCount）。</summary>
    public List<HorizontalAlign> ColumnAlignments { get; } = new();

    public int DataRowCount => Grid.Count;

    /// <summary>取某列的数值（跳过不可解析项），返回 (物理行号, 数值)。</summary>
    public List<(int RowNumber, double Value)> GetNumericColumn(int colIndex)
    {
        var list = new List<(int, double)>();
        for (int i = 0; i < Grid.Count; i++)
        {
            if (colIndex < 0 || colIndex >= Grid[i].Length) continue;
            var raw = Grid[i][colIndex];
            if (CellHelper.TryParseDouble(raw, out var v))
                list.Add((RowNumbers[i], v));
        }
        return list;
    }
}
