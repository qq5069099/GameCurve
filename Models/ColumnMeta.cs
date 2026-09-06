using GameCurve.Excel;

namespace GameCurve.Models;

/// <summary>
/// 某一列的元信息（来自第 1 行表头 + 数据样本）。
/// </summary>
public sealed class ColumnMeta
{
    /// <summary>0 基列索引。</summary>
    public int ColumnIndex { get; set; }

    /// <summary>列字母，如 "A"、"B"。 </summary>
    public string Letter { get; set; } = "";

    /// <summary>字段名（表头中第一个冒号前，可能带 * 等标记）。</summary>
    public string Name { get; set; } = "";

    /// <summary>中文/显示名（表头最后一段）。</summary>
    public string? Label { get; set; }

    /// <summary>类型标注，如 "long"、"int"、"Money[]"（<> 内的内容）。</summary>
    public string? Type { get; set; }

    /// <summary>原始表头文本。</summary>
    public string HeaderRaw { get; set; } = "";

    /// <summary>该列表头是否为空。</summary>
    public bool IsEmpty { get; set; }

    /// <summary>是否为可做曲线的标量数值列。</summary>
    public bool IsNumericScalar { get; set; }

    /// <summary>写回时是否按整数处理（long/int 等）。</summary>
    public bool IsInteger { get; set; }

    /// <summary>非空单元格数量。</summary>
    public int NonEmptyCount { get; set; }

    /// <summary>可解析为数值的单元格数量。</summary>
    public int NumericCount { get; set; }

    /// <summary>列内总数据行数。</summary>
    public int TotalRows { get; set; }

    /// <summary>展示列名：保留表头原始文本，不做简化。</summary>
    public string DisplayName =>
        IsEmpty || string.IsNullOrWhiteSpace(HeaderRaw)
            ? $"（空列 {Letter}）"
            : HeaderRaw.Trim();

    /// <summary>
    /// 根据表头文本刷新派生字段（名称/类型/标签/数值性）。用于新建列或重命名列后
    /// 同步元信息，无需重新扫描整表样本。dataRowCount 用于写入该列的 TotalRows。
    /// </summary>
    public void RefreshFromHeader(string headerRaw, int dataRowCount)
    {
        headerRaw = (headerRaw ?? "").Trim();
        HeaderRaw = headerRaw;
        Name = HeaderParser.ParseName(headerRaw);
        Label = HeaderParser.ParseLabel(headerRaw);
        Type = HeaderParser.ParseType(headerRaw);
        IsEmpty = string.IsNullOrWhiteSpace(headerRaw);
        IsNumericScalar = HeaderParser.IsExplicitNumericScalar(headerRaw, Type);
        IsInteger = HeaderParser.IsExplicitInteger(Type);
        TotalRows = dataRowCount;
    }

    public override string ToString() => Name == "" ? Letter : DisplayName;
}
