using System.Text.RegularExpressions;

namespace GameCurve.Excel;

/// <summary>
/// 解析表头单元格文本，提取字段名 / 类型标注 / 显示名，并判断是否为可编辑的标量数值列。
/// 表头形如：*v:&lt;long&gt;:数量、ID:序号、desc、price:&lt;Money&gt;.json:价格。
/// </summary>
public static class HeaderParser
{
    private static readonly HashSet<string> IntegerTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "long", "int", "short", "byte", "sbyte", "uint", "ulong", "ushort", "int32", "int64", "uint32", "uint64"
    };

    private static readonly HashSet<string> FloatTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "double", "float", "decimal", "single", "number", "num", "real", "numeric"
    };

    private static readonly HashSet<string> NumericTypes = new(IntegerTypes, StringComparer.OrdinalIgnoreCase);

    static HeaderParser()
    {
        NumericTypes.UnionWith(FloatTypes);
    }

    public static string ParseName(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return "";
        var first = raw.IndexOf(':');
        return (first >= 0 ? raw[..first] : raw).Trim();
    }

    public static string? ParseLabel(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        var parts = raw.Split(':', StringSplitOptions.TrimEntries);
        if (parts.Length < 2) return null;
        // 去掉最后的类型注解，取最后一段可读文本
        return parts[^1];
    }

    public static string? ParseType(string raw)
    {
        if (string.IsNullOrWhiteSpace(raw)) return null;
        var m = Regex.Match(raw, "<([^<>]+)>");
        return m.Success ? m.Groups[1].Value.Trim() : null;
    }

    /// <summary>表头是否明确标注为标量数值类型（非数组、非 json 复合对象）。</summary>
    public static bool IsExplicitNumericScalar(string raw, string? type)
    {
        if (string.IsNullOrWhiteSpace(type)) return false;
        if (type.Contains('[') || type.Contains(']')) return false;
        // 复合类型，如 <Money>.json / <Attr>.json
        if (raw.IndexOf(".json", StringComparison.OrdinalIgnoreCase) >= 0) return false;
        return NumericTypes.Contains(type);
    }

    public static bool IsExplicitInteger(string? type)
        => !string.IsNullOrWhiteSpace(type) && IntegerTypes.Contains(type);
}
