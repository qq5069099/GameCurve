using GameCurve.Excel;

namespace GameCurve.Models;

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
