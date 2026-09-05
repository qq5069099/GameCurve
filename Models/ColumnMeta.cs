namespace GameCurve.Models;

/// <summary>
/// 某一列的元信息（来自第 1 行表头 + 数据样本）。
/// </summary>
public sealed class ColumnMeta
{
    /// <summary>0 基列索引。</summary>
    public int ColumnIndex { get; init; }

    /// <summary>列字母，如 "A"、"B"。 </summary>
    public string Letter { get; init; } = "";

    /// <summary>字段名（表头中第一个冒号前，可能带 * 等标记）。</summary>
    public string Name { get; init; } = "";

    /// <summary>中文/显示名（表头最后一段）。</summary>
    public string? Label { get; set; }

    /// <summary>类型标注，如 "long"、"int"、"Money[]"（<> 内的内容）。</summary>
    public string? Type { get; init; }

    /// <summary>原始表头文本。</summary>
    public string HeaderRaw { get; init; } = "";

    /// <summary>该列表头是否为空。</summary>
    public bool IsEmpty { get; init; }

    /// <summary>是否为可做曲线的标量数值列。</summary>
    public bool IsNumericScalar { get; init; }

    /// <summary>写回时是否按整数处理（long/int 等）。</summary>
    public bool IsInteger { get; init; }

    /// <summary>非空单元格数量。</summary>
    public int NonEmptyCount { get; init; }

    /// <summary>可解析为数值的单元格数量。</summary>
    public int NumericCount { get; init; }

    /// <summary>列内总数据行数。</summary>
    public int TotalRows { get; init; }

    /// <summary>展示列名：保留表头原始文本，不做简化。</summary>
    public string DisplayName =>
        IsEmpty || string.IsNullOrWhiteSpace(HeaderRaw)
            ? $"（空列 {Letter}）"
            : HeaderRaw.Trim();

    public override string ToString() => Name == "" ? Letter : DisplayName;
}
