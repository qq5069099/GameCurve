namespace GameCurve.Models;

/// <summary>
/// 一条“子曲线”的来源：可对应整列标量，也可对应数组列中的某个元素，
/// 或 .json 列中某个含 v 的对象。
/// </summary>
public sealed class CurveColumnOption
{
    public ColumnMeta Column { get; init; } = null!;

    /// <summary>数组标量列中元素的下标；普通标量列为 -1。</summary>
    public int SubIndex { get; init; } = -1;

    /// <summary>.json 数组中的元素序号（无 ID 字段时用于定位）。</summary>
    public int JsonIndex { get; init; } = -1;

    /// <summary>.json 数组元素的 ID（有 ID 字段时优先按它定位）。</summary>
    public string JsonId { get; init; } = "";

    /// <summary>是否为 .json 的 v 子曲线。</summary>
    public bool IsJsonValue { get; init; }

    public string DisplayName { get; init; } = "";

    public bool IsInteger => Column.IsInteger;
    public bool IsSubCurve => SubIndex >= 0 || IsJsonValue;

    public override string ToString() => DisplayName;
}
