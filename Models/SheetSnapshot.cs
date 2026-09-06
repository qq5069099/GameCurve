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

    /// <summary>每列表头文字的显示对齐方式（按列索引，长度同 ColumnCount）。</summary>
    public List<HorizontalAlign> HeaderAlignments { get; } = new();

    public int DataRowCount => Grid.Count;

    /// <summary>
    /// 在物理行号 <paramref name="rowNumber"/> 处插入一个空行，其后的所有行物理行号整体 +1。
    /// 空行自动插入到网格的对应位置。
    /// </summary>
    public void InsertRow(int rowNumber)
    {
        int gridIndex = RowNumbers.FindIndex(r => r >= rowNumber);
        if (gridIndex < 0) gridIndex = RowNumbers.Count;
        for (int i = 0; i < RowNumbers.Count; i++)
            if (RowNumbers[i] >= rowNumber) RowNumbers[i] += 1;
        Grid.Insert(gridIndex, new string?[ColumnCount]);
        RowNumbers.Insert(gridIndex, rowNumber);
        if (rowNumber > MaxRow) MaxRow = rowNumber;
    }

    /// <summary>删除物理行号 <paramref name="rowNumber"/> 的行，其后的所有行物理行号整体 -1。</summary>
    public void DeleteRow(int rowNumber)
    {
        int gridIndex = RowNumbers.IndexOf(rowNumber);
        if (gridIndex < 0 || gridIndex >= Grid.Count) return;
        Grid.RemoveAt(gridIndex);
        RowNumbers.RemoveAt(gridIndex);
        for (int i = 0; i < RowNumbers.Count; i++)
            if (RowNumbers[i] > rowNumber) RowNumbers[i] -= 1;
    }

    /// <summary>
    /// 在物理列索引 <paramref name="colIndex"/> 处插入一列（其后各列索引 +1），
    /// 所有已有行在该位置补空单元格，并以 headerRaw 作为新列表头。
    /// </summary>
    public void InsertColumn(int colIndex, string headerRaw)
    {
        colIndex = Math.Clamp(colIndex, 0, ColumnCount);
        var meta = new ColumnMeta { ColumnIndex = colIndex, Letter = CellHelper.ColumnIndexToLetter(colIndex) };
        meta.RefreshFromHeader(headerRaw, DataRowCount);
        for (int i = colIndex; i < Columns.Count; i++)
        {
            Columns[i].ColumnIndex = i + 1;
            Columns[i].Letter = CellHelper.ColumnIndexToLetter(i + 1);
        }
        Columns.Insert(colIndex, meta);
        // 新加的列默认居中，方便直接输入内容（Excel 式）
        ColumnAlignments.Insert(colIndex, HorizontalAlign.Center);
        HeaderAlignments.Insert(colIndex, HorizontalAlign.Default);

        // 每行在 colIndex 处插入空单元格，列数 +1
        for (int gi = 0; gi < Grid.Count; gi++)
        {
            var old = Grid[gi];
            var newRow = new string?[ColumnCount + 1];
            for (int i = 0; i < old.Length; i++)
            {
                int j = i < colIndex ? i : i + 1;
                if (j < newRow.Length) newRow[j] = old[i];
            }
            Grid[gi] = newRow;
        }
        ColumnCount++;
    }

    /// <summary>删除物理列索引 <paramref name="colIndex"/> 处的列。</summary>
    public void DeleteColumn(int colIndex)
    {
        if (colIndex < 0 || colIndex >= Columns.Count) return;
        Columns.RemoveAt(colIndex);
        ColumnAlignments.RemoveAt(colIndex);
        HeaderAlignments.RemoveAt(colIndex);
        for (int i = colIndex; i < Columns.Count; i++)
        {
            Columns[i].ColumnIndex = i;
            Columns[i].Letter = CellHelper.ColumnIndexToLetter(i);
        }
        for (int gi = 0; gi < Grid.Count; gi++)
        {
            var old = Grid[gi];
            var newRow = new string?[ColumnCount - 1];
            for (int i = 0, j = 0; i < old.Length; i++)
            {
                if (i == colIndex) continue;
                if (j < newRow.Length) newRow[j++] = old[i];
            }
            Grid[gi] = newRow;
        }
        ColumnCount--;
    }

    /// <summary>重命名物理列索引 <paramref name="colIndex"/> 的表头。</summary>
    public void RenameColumn(int colIndex, string headerRaw)
    {
        if (colIndex < 0 || colIndex >= Columns.Count) return;
        Columns[colIndex].RefreshFromHeader(headerRaw, DataRowCount);
    }

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
