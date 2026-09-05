namespace GameCurve.Models;

/// <summary>
/// 曲线上一个数据点。默认 X = 物理行号，Y = 所选列单元格的值。
/// </summary>
public sealed class CurvePoint
{
    public double X { get; set; }
    public double Y { get; set; }

    /// <summary>该点所在 Excel 物理行号（1 基）。</summary>
    public int RowNumber { get; set; }

    /// <summary>X 是否绑定到 Excel 单元格（自定义 X 列时为真，行号模式为假）。</summary>
    public bool XEditable { get; set; }

    public CurvePoint() { }

    public CurvePoint(double x, double y, int rowNumber, bool xEditable)
    {
        X = x;
        Y = y;
        RowNumber = rowNumber;
        XEditable = xEditable;
    }

    public CurvePoint Clone() => new(X, Y, RowNumber, XEditable);
}

/// <summary>
/// 一个由某列生成的曲线序列。
/// </summary>
public sealed class CurveSeries
{
    public ColumnMeta Column { get; init; } = null!;
    public List<CurvePoint> Points { get; } = new();
    public Color Color { get; set; }
}
