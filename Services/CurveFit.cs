using System.Globalization;

namespace GameCurve.Services;

/// <summary>支持的拟合线型。</summary>
public enum FitModel
{
    Linear,      // y = a + b*x
    Polynomial,  // y = Σ c_i * x^i
    Exp2,        // y = a * e^(b*x)
    Exp3,        // y = a + b * e^(c*x)
    Logarithmic, // y = a + b*ln(x)
    Power,       // y = a * x^b
    Logistic,    // y = a + (b-a)/(1+e^(-k(x-m)))
    Gaussian,    // y = a + b*e^(-((x-c)^2)/(2*d^2))
    Sine,        // y = a + b*sin(c*x+d)
    Rational     // y = (a+b*x)/(1+c*x)
}

/// <summary>一次曲线拟合的结果：可求值函数、公式、拟合优度与收敛状态。</summary>
public sealed class FitResult
{
    public Func<double, double>? Evaluate { get; internal set; }
    public string Formula { get; internal set; } = "";
    public string Label { get; internal set; } = "";
    public double R2 { get; internal set; }
    public double RMSE { get; internal set; }
    public bool Converged { get; internal set; } = true;
    public string Error { get; internal set; } = "";
    public FitModel Model { get; internal set; }
}

/// <summary>
/// 对一组点做最小二乘曲线拟合。包含解析方法与 Levenberg-Marquardt 非线性求解，
/// 支持直线、任意次多项式、双/三参数指数、对数、幂、S 型、高斯、正弦、有理函数。
/// </summary>
public static class CurveFit
{
    public static string LabelOf(FitModel m) => m switch
    {
        FitModel.Linear => "直线",
        FitModel.Polynomial => "多项式",
        FitModel.Exp2 => "指数(双参)",
        FitModel.Exp3 => "指数(三参)",
        FitModel.Logarithmic => "对数",
        FitModel.Power => "幂函数",
        FitModel.Logistic => "S型(Logistic)",
        FitModel.Gaussian => "高斯",
        FitModel.Sine => "正弦",
        FitModel.Rational => "有理函数",
        _ => m.ToString()
    };

    public static int MaxPolynomialDegree => 8;

    /// <summary>对点列做拟合（点会按 X 升序排序后处理）。degree 仅对 Polynomial 生效。</summary>
    public static FitResult Fit(IReadOnlyList<(double X, double Y)> pts, FitModel model, int degree = 2)
    {
        var data = pts.Where(p => double.IsFinite(p.X) && double.IsFinite(p.Y)).OrderBy(p => p.X).ToList();
        var result = new FitResult { Model = model, Label = LabelOf(model) };
        if (data.Count < 2)
        {
            result.Error = "有效拟合点不足，至少需要 2 个点";
            return result;
        }

        switch (model)
        {
            case FitModel.Linear:
                FitLinear(data, result);
                break;
            case FitModel.Polynomial:
                FitPolynomial(data, Math.Clamp(degree, 2, MaxPolynomialDegree), result);
                break;
            case FitModel.Exp2:
                FitExp2(data, result);
                break;
            case FitModel.Exp3:
                FitNonlinear(data, model, result);
                break;
            case FitModel.Logarithmic:
                FitLogarithmic(data, result);
                break;
            case FitModel.Power:
                FitPower(data, result);
                break;
            case FitModel.Logistic:
            case FitModel.Gaussian:
            case FitModel.Sine:
            case FitModel.Rational:
                FitNonlinear(data, model, result);
                break;
            default:
                result.Error = "未知拟合类型";
                break;
        }

        if (result.Evaluate != null)
        {
            var (r2, rmse) = Goodness(data, result.Evaluate);
            result.R2 = r2;
            result.RMSE = rmse;
        }
        return result;
    }

    // ---------- 解析方法 ----------
    private static void FitLinear(List<(double X, double Y)> d, FitResult r)
    {
        var (a, b) = FitLine(d);
        r.Evaluate = x => a + b * x;
        r.Formula = $"y = {F(a)} {(b >= 0 ? "+" : "-")} {F(Math.Abs(b))}x";
        r.Converged = true;
    }

    private static void FitExp2(List<(double X, double Y)> d, FitResult r)
    {
        if (d.Any(p => p.Y <= 0))
        {
            r.Error = "双参指数要求所有 Y 均大于 0";
            return;
        }
        var t = d.Select(p => (p.X, Math.Log(p.Y))).ToList();
        var (a, b) = FitLine(t);
        double amp = Math.Exp(a);
        r.Evaluate = x => amp * Math.Exp(b * x);
        r.Formula = $"y = {F(amp)} * e^({F(b)}x)";
        r.Converged = true;
    }

    private static void FitLogarithmic(List<(double X, double Y)> d, FitResult r)
    {
        if (d.Any(p => p.X <= 0))
        {
            r.Error = "对数曲线要求所有 X 均大于 0";
            return;
        }
        var t = d.Select(p => (Math.Log(p.X), p.Y)).ToList();
        var (a, b) = FitLine(t);
        r.Evaluate = x => a + b * Math.Log(x);
        r.Formula = $"y = {F(a)} {(b >= 0 ? "+" : "-")} {F(Math.Abs(b))}*ln(x)";
        r.Converged = true;
    }

    private static void FitPower(List<(double X, double Y)> d, FitResult r)
    {
        if (d.Any(p => p.X <= 0 || p.Y <= 0))
        {
            r.Error = "幂函数要求所有 X 与 Y 均大于 0";
            return;
        }
        var t = d.Select(p => (Math.Log(p.X), Math.Log(p.Y))).ToList();
        var (a, b) = FitLine(t);
        double amp = Math.Exp(a);
        r.Evaluate = x => amp * Math.Pow(x, b);
        r.Formula = $"y = {F(amp)} * x^{F(b)}";
        r.Converged = true;
    }

    private static void FitPolynomial(List<(double X, double Y)> d, int degree, FitResult r)
    {
        int m = degree + 1;
        if (d.Count <= degree)
        {
            r.Error = $"{degree} 次多项式至少需要 {m} 个点";
            return;
        }

        var mat = new double[m, m];
        var rhs = new double[m];
        for (int i = 0; i < m; i++)
            for (int j = 0; j < m; j++)
            {
                double s = 0;
                foreach (var p in d) s += Math.Pow(p.X, i + j);
                mat[i, j] = s;
            }
        for (int i = 0; i < m; i++)
        {
            double s = 0;
            foreach (var p in d) s += p.Y * Math.Pow(p.X, i);
            rhs[i] = s;
        }

        if (!SolveLinear(mat, rhs, out var coeff))
        {
            r.Error = "数据点过于集中，无法拟合该多项式";
            return;
        }
        r.Evaluate = x =>
        {
            double s = 0;
            for (int i = coeff.Length - 1; i >= 0; i--) s = s * x + coeff[i];
            return s;
        };
        r.Formula = BuildPolyFormula(coeff);
        r.Converged = true;
    }

    // ---------- 非线性（Levenberg-Marquardt） ----------
    private static void FitNonlinear(List<(double X, double Y)> d, FitModel model, FitResult r)
    {
        double[] p0;
        Func<double, double[], double> modelFunc;
        switch (model)
        {
            case FitModel.Exp3:
                p0 = InitExp3(d);
                modelFunc = (x, p) => p[0] + p[1] * Math.Exp(p[2] * x);
                break;
            case FitModel.Logistic:
                p0 = InitLogistic(d);
                modelFunc = (x, p) => p[0] + (p[1] - p[0]) / (1.0 + Math.Exp(-p[2] * (x - p[3])));
                break;
            case FitModel.Gaussian:
                p0 = InitGaussian(d);
                modelFunc = (x, p) => p[0] + p[1] * Math.Exp(-(x - p[2]) * (x - p[2]) / (2.0 * p[3] * p[3]));
                break;
            case FitModel.Sine:
                p0 = InitSine(d);
                modelFunc = (x, p) => p[0] + p[1] * Math.Sin(p[2] * x + p[3]);
                break;
            case FitModel.Rational:
                p0 = InitRational(d);
                modelFunc = (x, p) => (p[0] + p[1] * x) / (1.0 + p[2] * x);
                break;
            default:
                r.Error = "不支持的拟合类型";
                return;
        }

        if (!LevenbergMarquardt(d, p0, modelFunc, out var p, out var lmErr))
        {
            r.Error = lmErr;
            return;
        }
        double finalSse = Sse(d, p, modelFunc);
        if (!double.IsFinite(finalSse))
        {
            r.Error = "拟合不收敛，请尝试其他线型或减少数据范围";
            return;
        }

        r.Evaluate = x => modelFunc(x, p);
        r.Formula = BuildNonlinearFormula(model, p);
        r.Converged = true;
    }

    private static double[] InitExp3(List<(double X, double Y)> d)
    {
        double minX = d[0].X, maxX = d[^1].X;
        double minY = d.Min(p => p.Y), maxY = d.Max(p => p.Y);
        bool up = d[^1].Y >= d[0].Y;
        double range = Math.Max(1e-6, maxX - minX);
        double a0 = up ? minY : maxY;
        double b0 = (up ? 1.0 : -1.0) * Math.Max(1e-6, Math.Abs(maxY - minY));
        double c0 = (up ? 1.0 : -1.0) * Math.Log(2.0) / (range * 0.5);
        int i1 = Math.Max(0, d.Count / 4), i2 = Math.Min(d.Count - 1, d.Count * 3 / 4);
        double y1 = d[i1].Y - a0, y2 = d[i2].Y - a0;
        if (y1 > 1e-12 && y2 > 1e-12 && Math.Abs(d[i2].X - d[i1].X) > 1e-9)
            c0 = Math.Log(y2 / y1) / (d[i2].X - d[i1].X);
        return new[] { a0, b0, c0 };
    }

    private static double[] InitLogistic(List<(double X, double Y)> d)
    {
        double minX = d[0].X, maxX = d[^1].X;
        double minY = d.Min(p => p.Y), maxY = d.Max(p => p.Y);
        bool up = d[^1].Y >= d[0].Y;
        double range = Math.Max(1e-6, maxX - minX);
        double lo = minY, hi = maxY;
        if (Math.Abs(hi - lo) < 1e-9) hi = lo + 1;
        double mid = 0.5 * (minX + maxX);
        double k0 = 4.0 / range;
        int i0 = Math.Max(0, (int)Math.Floor(mid) - 1);
        int i1 = Math.Min(d.Count - 1, (int)Math.Floor(mid) + 1);
        if (i1 > i0 && Math.Abs(d[i1].X - d[i0].X) > 1e-9)
        {
            double slope = (d[i1].Y - d[i0].Y) / (d[i1].X - d[i0].X);
            if (Math.Abs(slope) > 1e-12 && Math.Abs(hi - lo) > 1e-9)
                k0 = 4.0 * Math.Abs(slope) / Math.Abs(hi - lo);
        }
        k0 = Math.Max(1e-6, Math.Abs(k0)) * (up ? 1.0 : -1.0);
        return new[] { lo, hi, k0, mid };
    }

    private static double[] InitGaussian(List<(double X, double Y)> d)
    {
        double minY = d.Min(p => p.Y), maxY = d.Max(p => p.Y);
        double den = d.Sum(p => p.Y);
        double c0 = Math.Abs(den) < 1e-12 ? d[0].X : d.Sum(p => p.X * p.Y) / den;
        double range = d[^1].X - d[0].X;
        double d0 = Math.Max(range / 5.0, 1e-6);
        return new[] { minY, Math.Max(1e-6, maxY - minY), c0, d0 };
    }

    private static double[] InitSine(List<(double X, double Y)> d)
    {
        double meanY = d.Average(p => p.Y);
        double minY = d.Min(p => p.Y), maxY = d.Max(p => p.Y);
        double range = Math.Max(1e-6, d[^1].X - d[0].X);
        return new[] { meanY, Math.Max(1e-6, (maxY - minY) / 2.0), 2.0 * Math.PI / range, 0.0 };
    }

    private static double[] InitRational(List<(double X, double Y)> d)
    {
        double meanY = d.Average(p => p.Y);
        double range = Math.Max(1e-6, d[^1].X - d[0].X);
        double spread = d[^1].Y - d[0].Y;
        return new[] { meanY, spread / range, 0.0 };
    }

    private static string BuildNonlinearFormula(FitModel model, double[] p) => model switch
    {
        FitModel.Exp3 => $"y = {F(p[0])} + {F(p[1])}*e^({F(p[2])}x)",
        FitModel.Logistic => $"y = {F(p[0])} + ({F(p[1] - p[0])})/(1+e^(-{F(p[2])}(x-{F(p[3])})))",
        FitModel.Gaussian => $"y = {F(p[0])} + {F(p[1])}*e^(-((x-{F(p[2])})^2)/{F(2.0 * p[3] * p[3])})",
        FitModel.Sine => $"y = {F(p[0])} + {F(p[1])}*sin({F(p[2])}x+{F(p[3])})",
        FitModel.Rational => $"y = ({F(p[0])}+{F(p[1])}x)/(1+{F(p[2])}x)",
        _ => ""
    };

    /// <summary>Levenberg-Marquardt：数值雅可比 + 阻尼高斯-牛顿。</summary>
    private static bool LevenbergMarquardt(
        List<(double X, double Y)> d,
        double[] p0,
        Func<double, double[], double> model,
        out double[] p,
        out string error)
    {
        error = "";
        int n = d.Count, m = p0.Length;
        if (n < m)
        {
            error = $"该线型需要至少 {m} 个点";
            p = p0;
            return false;
        }
        p = (double[])p0.Clone();
        double lambda = 1e-3;
        double current = Sse(d, p, model);
        if (!double.IsFinite(current)) current = double.MaxValue / 2;

        for (int iter = 0; iter < 120; iter++)
        {
            if (lambda > 1e12)
            {
                error = "未收敛（阻尼过大），请换线型或检查数据";
                return false;
            }

            var jac = new double[n, m];
            for (int i = 0; i < n; i++)
            {
                double xi = d[i].X;
                for (int j = 0; j < m; j++)
                {
                    double h = 1e-6 * Math.Max(1.0, Math.Abs(p[j]));
                    var qp = (double[])p.Clone(); qp[j] += h;
                    var qm = (double[])p.Clone(); qm[j] -= h;
                    double fp = model(xi, qp), fm = model(xi, qm);
                    if (double.IsFinite(fp) && double.IsFinite(fm))
                        jac[i, j] = (fp - fm) / (2.0 * h);
                    else
                        jac[i, j] = 0;
                }
            }

            var a = new double[m, m];
            var g = new double[m];
            for (int i = 0; i < n; i++)
            {
                double yi = d[i].Y;
                double ri = yi - model(d[i].X, p);
                if (!double.IsFinite(ri)) ri = 0;
                for (int j = 0; j < m; j++)
                {
                    g[j] += jac[i, j] * ri;
                    for (int k = j; k < m; k++)
                        a[j, k] += jac[i, j] * jac[i, k];
                }
            }
            for (int j = 0; j < m; j++)
            {
                a[j, j] = a[j, j] * (1.0 + lambda) + 1e-12;
                for (int k = j + 1; k < m; k++) a[k, j] = a[j, k];
            }

            if (!SolveLinear(a, g, out var delta))
            {
                lambda *= 10.0;
                continue;
            }

            var candidate = (double[])p.Clone();
            for (int j = 0; j < m; j++) candidate[j] += delta[j];
            double next = Sse(d, candidate, model);
            if (double.IsFinite(next) && next < current)
            {
                double improvement = Math.Abs(current - next) / Math.Max(1.0, current);
                p = candidate;
                current = next;
                lambda = Math.Max(lambda / 3.0, 1e-12);
                if (improvement < 1e-10) break;
            }
            else
            {
                lambda *= 5.0;
            }
        }
        return true;
    }

    private static double Sse(List<(double X, double Y)> d, double[] p, Func<double, double[], double> model)
    {
        double s = 0;
        foreach (var pt in d)
        {
            double f = model(pt.X, p);
            if (!double.IsFinite(f)) return double.MaxValue;
            double e = pt.Y - f;
            s += e * e;
            if (s > 1e30) return double.MaxValue;
        }
        return s;
    }

    private static (double R2, double RMSE) Goodness(List<(double X, double Y)> d, Func<double, double> f)
    {
        double meanY = d.Average(p => p.Y);
        double ssRes = 0, ssTot = 0;
        foreach (var p in d)
        {
            double e = p.Y - f(p.X);
            ssRes += e * e;
            ssTot += (p.Y - meanY) * (p.Y - meanY);
        }
        double r2 = ssTot < 1e-12 ? (ssRes < 1e-12 ? 1 : 0) : 1 - ssRes / ssTot;
        if (r2 < 0) r2 = 0;
        double rmse = Math.Sqrt(ssRes / Math.Max(1, d.Count));
        return (r2, rmse);
    }

    // ---------- 工具 ----------
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
            return (sy / n, 0);
        double b = (n * sxy - sx * sy) / denom;
        double a = (sy - b * sx) / n;
        return (a, b);
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
