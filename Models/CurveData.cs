namespace GameCurve.Models;

/// <summary>
/// 曲线上一个数据点。“行号”X 模式下 X 使用排除表头行后的数据行序号（1 基），
/// Y 为所选列单元格的值；RowNumber 始终保留 Excel 物理行号，供编辑映射使用。
/// </summary>
public sealed class CurvePoint
{
    public double X { get; set; }
    public double Y { get; set; }

    /// <summary>该点所在 Excel 物理行号（1 基）。</summary>
    public int RowNumber { get; set; }

    /// <summary>排除表头行后的数据行序号（1 基），与表格“行号”列一致。</summary>
    public int DataRow { get; set; }

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

    public CurvePoint Clone() => new(X, Y, RowNumber, XEditable) { DataRow = DataRow };
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
