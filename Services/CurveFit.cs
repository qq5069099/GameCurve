using System.Globalization;

namespace GameCurve.Services;

/// <summary>
/// 对一组点做最小二乘曲线拟合，并给出可求值的函数与公式文本。
/// 支持直线、指数、对数、幂、二次、三次多项式。
/// </summary>
public static class CurveFit
{
    public enum Kind
    {
        Linear,        // y = a + b*x
        Exponential,   // y = a * e^(b*x)
        Logarithmic,   // y = a + b*ln(x)
        Power,         // y = a * x^b
        Quadratic,     // y = c0 + c1*x + c2*x^2
        Cubic          // y = c0 + c1*x + c2*x^2 + c3*x^3
    }

    public static string NameOf(Kind k) => k switch
    {
        Kind.Linear => "直线",
        Kind.Exponential => "指数曲线",
        Kind.Logarithmic => "对数曲线",
        Kind.Power => "幂函数",
        Kind.Quadratic => "二次多项式",
        Kind.Cubic => "三次多项式",
        _ => k.ToString()
    };

    /// <summary>
    /// 对点列做拟合。成功返回求值函数，否则返回 null 并给出 error。
    /// </summary>
    public static Func<double, double>? Fit(
        IReadOnlyList<(double X, double Y)> pts,
        Kind kind,
        out string formula,
        out string error)
    {
        formula = "";
        error = "";
        var data = pts.Where(p => double.IsFinite(p.X) && double.IsFinite(p.Y)).ToList();
        if (data.Count < 2)
        {
            error = "有效拟合点不足，至少需要 2 个点";
            return null;
        }

        switch (kind)
        {
            case Kind.Linear:
                return FitLinear(data, out formula);
            case Kind.Exponential:
                return FitExponential(data, out formula, out error);
            case Kind.Logarithmic:
                return FitLogarithmic(data, out formula, out error);
            case Kind.Power:
                return FitPower(data, out formula, out error);
            case Kind.Quadratic:
                return FitPolynomial(data, 2, out formula, out error);
            case Kind.Cubic:
                return FitPolynomial(data, 3, out formula, out error);
            default:
                error = "未知拟合类型";
                return null;
        }
    }

    /// <summary>普通最小二乘直线拟合，返回 (截距 a, 斜率 b)。</summary>
    private static (double A, double B) FitLine(IReadOnlyList<(double X, double Y)> d)
    {
        int n = d.Count;
        double sx = 0, sy = 0, sxx = 0, sxy = 0;
        foreach (var p in d)
        {
            sx += p.X; sy += p.Y; sxx += p.X * p.X; sxy += p.X * p.Y;
        }
        double denom = n * sxx - sx * sx;
        if (Math.Abs(denom) < 1e-12)
            return (sy / n, 0); // X 太集中，退化为水平线
        double b = (n * sxy - sx * sy) / denom;
        double a = (sy - b * sx) / n;
        return (a, b);
    }

    private static Func<double, double> FitLinear(List<(double X, double Y)> d, out string formula)
    {
        var (a, b) = FitLine(d);
        formula = $"y = {F(a)} {(b >= 0 ? "+" : "-")} {F(Math.Abs(b))}x";
        return x => a + b * x;
    }

    private static Func<double, double>? FitExponential(List<(double X, double Y)> d, out string formula, out string error)
    {
        formula = "";
        if (d.Any(p => p.Y <= 0))
        {
            error = "指数曲线要求所有 Y 均大于 0";
            return null;
        }
        var t = d.Select(p => (p.X, Math.Log(p.Y))).ToList();
        var (a, b) = FitLine(t);
        double amp = Math.Exp(a);
        formula = $"y = {F(amp)} * e^({F(b)}x)";
        error = "";
        return x => amp * Math.Exp(b * x);
    }

    private static Func<double, double>? FitLogarithmic(List<(double X, double Y)> d, out string formula, out string error)
    {
        formula = "";
        if (d.Any(p => p.X <= 0))
        {
            error = "对数曲线要求所有 X 均大于 0";
            return null;
        }
        var t = d.Select(p => (Math.Log(p.X), p.Y)).ToList();
        var (a, b) = FitLine(t);
        formula = $"y = {F(a)} {(b >= 0 ? "+" : "-")} {F(Math.Abs(b))}*ln(x)";
        error = "";
        return x => a + b * Math.Log(x);
    }

    private static Func<double, double>? FitPower(List<(double X, double Y)> d, out string formula, out string error)
    {
        formula = "";
        if (d.Any(p => p.X <= 0 || p.Y <= 0))
        {
            error = "幂函数要求所有 X 与 Y 均大于 0";
            return null;
        }
        var t = d.Select(p => (Math.Log(p.X), Math.Log(p.Y))).ToList();
        var (a, b) = FitLine(t);
        double amp = Math.Exp(a);
        formula = $"y = {F(amp)} * x^{F(b)}";
        error = "";
        return x => amp * Math.Pow(x, b);
    }

    /// <summary>通用多项式最小二乘（正规方程 + 列主元高斯消元）。</summary>
    private static Func<double, double>? FitPolynomial(List<(double X, double Y)> d, int degree, out string formula, out string error)
    {
        formula = "";
        int m = degree + 1;
        if (d.Count <= degree)
        {
            error = $"{NameOf(degree == 2 ? Kind.Quadratic : Kind.Cubic)}拟合至少需要 {m} 个点";
            return null;
        }

        var mat = new double[m, m];
        var rhs = new double[m];
        for (int r = 0; r < m; r++)
            for (int c = 0; c < m; c++)
            {
                double s = 0;
                foreach (var p in d) s += Math.Pow(p.X, r + c);
                mat[r, c] = s;
            }
        for (int r = 0; r < m; r++)
        {
            double s = 0;
            foreach (var p in d) s += p.Y * Math.Pow(p.X, r);
            rhs[r] = s;
        }

        if (!SolveLinear(mat, rhs, out var coeff))
        {
            error = "数据点过于集中，无法拟合该多项式";
            return null;
        }

        error = "";
        formula = BuildPolyFormula(coeff);
        return x =>
        {
            double s = 0;
            for (int i = coeff.Length - 1; i >= 0; i--) s = s * x + coeff[i];
            return s;
        };
    }

    private static string BuildPolyFormula(double[] c)
    {
        var sb = new System.Text.StringBuilder("y = ");
        bool any = false;
        for (int i = 0; i < c.Length; i++)
        {
            double v = Math.Abs(c[i]);
            if (v < 1e-10) continue;
            if (any) sb.Append(c[i] >= 0 ? " + " : " - ");
            else if (c[i] < 0) sb.Append("-");
            if (i == 0) sb.Append(F(v));
            else if (i == 1) sb.Append(F(v) == "1" ? "x" : $"{F(v)}x");
            else sb.Append(F(v) == "1" ? $"x^{i}" : $"{F(v)}x^{i}");
            any = true;
        }
        if (!any) sb.Append("0");
        return sb.ToString();
    }

    /// <summary>解线性方程组 Ax=b，列主元高斯消元；奇异返回 false。</summary>
    private static bool SolveLinear(double[,] a, double[] b, out double[] x)
    {
        int n = b.Length;
        x = new double[n];
        var m = new double[n, n + 1];
        for (int i = 0; i < n; i++)
        {
            for (int j = 0; j < n; j++) m[i, j] = a[i, j];
            m[i, n] = b[i];
        }

        for (int col = 0; col < n; col++)
        {
            int piv = col;
            for (int r = col + 1; r < n; r++)
                if (Math.Abs(m[r, col]) > Math.Abs(m[piv, col])) piv = r;
            if (Math.Abs(m[piv, col]) < 1e-12)
            {
                x = Array.Empty<double>();
                return false;
            }
            if (piv != col)
                for (int j = 0; j <= n; j++)
                    (m[col, j], m[piv, j]) = (m[piv, j], m[col, j]);

            double d = m[col, col];
            for (int j = col; j <= n; j++) m[col, j] /= d;
            for (int r = 0; r < n; r++)
            {
                if (r == col) continue;
                double factor = m[r, col];
                if (Math.Abs(factor) < 1e-14) continue;
                for (int j = col; j <= n; j++) m[r, j] -= factor * m[col, j];
            }
        }
        for (int i = 0; i < n; i++) x[i] = m[i, n];
        return true;
    }

    private static string F(double v)
    {
        if (Math.Abs(v) < 1e-10) v = 0;
        return v.ToString("0.####", CultureInfo.InvariantCulture);
    }
}
