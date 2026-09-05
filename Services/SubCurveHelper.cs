using System.Globalization;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using GameCurve.Excel;
using GameCurve.Models;

namespace GameCurve.Services;

/// <summary>
/// 把“子曲线”从 Excel 单元格字符串里解析出来，并在修改后重新生成整格文本。
/// 支持三类：
/// 1) &lt;float[N]&gt; 数组标量列，每个元素一条曲线；
/// 2) .json 数组（如 Attr[]），每个含 v 的对象一条曲线；
/// 3) .json 对象含 v，一条曲线。
/// </summary>
public static class SubCurveHelper
{
    private static readonly Regex ArrayType = new(
        @"^(?<type>[A-Za-z]+)\[(?<len>\d*)\]$",
        RegexOptions.Compiled | RegexOptions.CultureInvariant);

    private static readonly HashSet<string> NumericArrayTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "float", "double", "int", "long", "short", "byte", "sbyte",
        "uint", "ulong", "ushort", "decimal"
    };

    private static readonly HashSet<string> IntegerArrayTypes = new(StringComparer.OrdinalIgnoreCase)
    {
        "int", "long", "short", "byte", "sbyte",
        "uint", "ulong", "ushort", "int32", "int64", "uint32", "uint64"
    };

    public static bool IsArrayCol(ColumnMeta col)
    {
        if (IsRangeColumn(col)) return true;
        if (string.IsNullOrWhiteSpace(col.Type)) return false;
        var m = ArrayType.Match(col.Type!);
        return m.Success && NumericArrayTypes.Contains(m.Groups["type"].Value);
    }

    /// <summary>识别 &lt;long&gt;.range / &lt;float&gt;.range 之类的区域列。</summary>
    public static bool IsRangeColumn(ColumnMeta col)
        => !string.IsNullOrWhiteSpace(col.HeaderRaw)
            && col.HeaderRaw.Contains(".range", StringComparison.OrdinalIgnoreCase)
            && !string.IsNullOrWhiteSpace(col.Type)
            && NumericArrayTypes.Contains(col.Type);

    /// <summary>判断数组列的元素基础类型是否为整型（int/long 等）。</summary>
    public static bool IsIntegerArray(ColumnMeta col)
    {
        if (string.IsNullOrWhiteSpace(col.Type)) return false;
        if (IsRangeColumn(col)) return IntegerArrayTypes.Contains(col.Type);
        var m = ArrayType.Match(col.Type!);
        return m.Success && IntegerArrayTypes.Contains(m.Groups["type"].Value);
    }

    public static bool IsJsonCol(ColumnMeta col)
        => !string.IsNullOrWhiteSpace(col.HeaderRaw)
            && col.HeaderRaw.Contains(".json", StringComparison.OrdinalIgnoreCase);

    public static int ArrayLength(ColumnMeta col)
    {
        if (!IsArrayCol(col)) return 0;
        var m = ArrayType.Match(col.Type!);
        return m.Success && int.TryParse(m.Groups["len"].Value, out var n) ? n : 0;
    }

    /// <summary>生成某表所有可显示的子曲线选项（含标量列）。</summary>
    public static List<CurveColumnOption> BuildOptions(SheetSnapshot snap)
    {
        var list = new List<CurveColumnOption>();
        foreach (var col in snap.Columns)
        {
            if (IsArrayCol(col))
            {
                // 长度不固定，按整列实际数据动态扫描
                int n = ScanArrayCount(snap, col);
                for (int i = 0; i < n; i++)
                    list.Add(new CurveColumnOption
                    {
                        Column = col,
                        SubIndex = i,
                        DisplayName = $"{ShortName(col)}[{i}]"
                    });
            }
            else if (col.IsNumericScalar)
            {
                list.Add(new CurveColumnOption
                {
                    Column = col,
                    DisplayName = col.DisplayName
                });
            }
            else if (IsJsonCol(col))
            {
                AddJsonOptions(snap, col, list);
            }
        }
        return list;
    }

    private static void AddJsonOptions(SheetSnapshot snap, ColumnMeta col, List<CurveColumnOption> list)
    {
        // 收集出现过的“含 v 的对象”键；优先用对象 ID，否则用数组序号
        var keys = new Dictionary<string, (string Key, int Index, bool HasId)>();
        int ctx = 0;
        for (int gi = 0; gi < snap.Grid.Count; gi++)
        {
            var raw = CellText(snap, gi, col.ColumnIndex);
            if (string.IsNullOrWhiteSpace(raw)) continue;
            try
            {
                using var doc = JsonDocument.Parse(raw);
                var root = doc.RootElement;
                if (root.ValueKind == JsonValueKind.Array)
                {
                    int idx = 0;
                    foreach (var el in root.EnumerateArray())
                    {
                        if (el.ValueKind == JsonValueKind.Object &&
                            el.TryGetProperty("v", out var v) && v.ValueKind == JsonValueKind.Number)
                        {
                            string id = "";
                            bool hasId = el.TryGetProperty("ID", out var idv) && idv.ValueKind == JsonValueKind.Number;
                            if (hasId) id = idv.GetDouble().ToString(CultureInfo.InvariantCulture);
                            string key = hasId ? id : idx.ToString(CultureInfo.InvariantCulture);
                            keys.TryAdd(key, (id, idx, hasId));
                        }
                        idx++;
                    }
                }
                else if (root.ValueKind == JsonValueKind.Object &&
                         root.TryGetProperty("v", out var v) && v.ValueKind == JsonValueKind.Number)
                {
                    keys.TryAdd("v", ("v", 0, false));
                }
            }
            catch { }
            if (++ctx > 20000) break;
        }

        foreach (var (key, meta) in keys)
        {
            var suffix = meta.HasId ? $"ID={meta.Key}" : (key == "v" ? "" : $"#{meta.Index}");
            list.Add(new CurveColumnOption
            {
                Column = col,
                IsJsonValue = true,
                JsonIndex = meta.Index,
                JsonId = meta.HasId ? meta.Key : "",
                DisplayName = string.IsNullOrEmpty(suffix) ? ShortName(col) : $"{ShortName(col)}[{suffix}]"
            });
        }
    }

    private static string CellText(SheetSnapshot snap, int gi, int col)
    {
        if (gi < 0 || gi >= snap.Grid.Count) return "";
        var row = snap.Grid[gi];
        return col >= 0 && col < row.Length ? row[col] ?? "" : "";
    }

    public static string ShortName(ColumnMeta col)
    {
        string label = col.Label ?? col.Name;
        if (string.IsNullOrWhiteSpace(label)) label = col.Letter;
        return label.Trim();
    }

    /// <summary>从某行单元格里提取子曲线值。</summary>
    public static bool TryReadValue(SheetSnapshot? snap, CurveColumnOption opt, int gi, out double value)
    {
        value = 0;
        if (snap == null) return false;
        var raw = CellText(snap, gi, opt.Column.ColumnIndex);
        if (string.IsNullOrWhiteSpace(raw)) return false;

        if (!opt.IsSubCurve)
            return CellHelper.TryParseDouble(raw, out value);

        if (opt.SubIndex >= 0)
            return TryArrayRead(raw, opt.SubIndex, out value);

        try
        {
            using var doc = JsonDocument.Parse(raw);
            var root = doc.RootElement;
            if (root.ValueKind == JsonValueKind.Array)
            {
                if (opt.JsonId.Length > 0)
                {
                    foreach (var el in root.EnumerateArray())
                        if (el.ValueKind == JsonValueKind.Object &&
                            el.TryGetProperty("ID", out var idv) &&
                            idv.ValueKind == JsonValueKind.Number &&
                            idv.GetDouble().ToString(CultureInfo.InvariantCulture) == opt.JsonId &&
                            el.TryGetProperty("v", out var vv))
                        {
                            return TryNumber(vv, out value);
                        }
                    return false;
                }

                int idx = 0;
                foreach (var el in root.EnumerateArray())
                {
                    if (idx == opt.JsonIndex && el.ValueKind == JsonValueKind.Object &&
                        el.TryGetProperty("v", out var vArr))
                        return TryNumber(vArr, out value);
                    idx++;
                }
                return false;
            }

            if (root.ValueKind == JsonValueKind.Object && root.TryGetProperty("v", out var vObj))
                return TryNumber(vObj, out value);
            return false;
        }
        catch { return false; }
    }

    private static bool TryArrayRead(string raw, int index, out double value)
    {
        value = 0;
        var parts = SplitArray(raw);
        if (index < 0 || index >= parts.Length) return false;
        return CellHelper.TryParseDouble(parts[index], out value);
    }

    private static string[] SplitArray(string raw)
    {
        var inner = raw.Trim();
        if (inner.StartsWith('[')) inner = inner[1..];
        if (inner.EndsWith(']')) inner = inner[..^1];
        return inner.Split(',', StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
    }

    private static int ScanArrayCount(SheetSnapshot snap, ColumnMeta col)
    {
        int max = 0;
        for (int gi = 0; gi < snap.Grid.Count; gi++)
        {
            var raw = CellText(snap, gi, col.ColumnIndex);
            if (string.IsNullOrWhiteSpace(raw)) continue;
            max = Math.Max(max, SplitArray(raw).Length);
        }
        return max;
    }

    private static bool TryNumber(JsonElement el, out double value)
    {
        value = 0;
        if (el.ValueKind == JsonValueKind.Number && el.TryGetDouble(out value)) return true;
        if (el.ValueKind == JsonValueKind.String && double.TryParse(el.GetString(), NumberStyles.Any, CultureInfo.InvariantCulture, out value)) return true;
        return false;
    }

    /// <summary>修改子曲线后，重新生成整格文本（数组或 JSON 对象字符串）。</summary>
    public static string SetValue(string raw, CurveColumnOption opt, double y)
    {
        bool asLong = IsIntegerArray(opt.Column) || opt.IsJsonValue;
        if (asLong) y = Math.Round(y, MidpointRounding.AwayFromZero);
        if (opt.SubIndex >= 0)
            return SetArrayValue(raw, opt.SubIndex, y);

        try
        {
            var node = JsonNode.Parse(raw);
            if (node == null) return raw;
            JsonObject? target = null;
            if (node is JsonArray arr)
            {
                if (opt.JsonId.Length > 0)
                {
                    foreach (var el in arr)
                        if (el is JsonObject obj &&
                            obj["ID"] is JsonValue idv &&
                            idv.GetValue<double>().ToString(CultureInfo.InvariantCulture) == opt.JsonId)
                        { target = obj; break; }
                }
                else if (opt.JsonIndex >= 0 && opt.JsonIndex < arr.Count && arr[opt.JsonIndex] is JsonObject o)
                    target = o;
            }
            else if (node is JsonObject obj)
                target = obj;

            if (target != null)
                target["v"] = asLong ? JsonValue.Create((long)y) : JsonValue.Create(y);
            return node.ToJsonString(new JsonSerializerOptions
            {
                WriteIndented = false,
                NumberHandling = System.Text.Json.Serialization.JsonNumberHandling.Strict
            });
        }
        catch { return raw; }
    }

    private static string SetArrayValue(string raw, int index, double y)
    {
        var parts = SplitArray(raw).ToList();
        if (index < 0 || index >= parts.Count) return raw;
        parts[index] = FormatNum(y);
        return "[" + string.Join(",", parts) + "]";
    }

    public static string FormatNum(double v)
    {
        if (Math.Abs(v - Math.Round(v)) < 1e-9) return Math.Round(v).ToString("0", CultureInfo.InvariantCulture);
        return v.ToString("0.######", CultureInfo.InvariantCulture);
    }
}
