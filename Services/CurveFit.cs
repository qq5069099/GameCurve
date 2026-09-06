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

/// <summary>一个拟合参数的名称与取值。</summary>
public sealed class FitParameter
{
    public string Name { get; }
    public double Value { get; }
    public FitParameter(string name, double value) { Name = name; Value = value; }
}

/// <summary>一次曲线拟合的结果：可求值函数、公式、拟合优度与参数。</summary>
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
    public IReadOnlyList<FitParameter> Parameters { get; internal set; } = Array.Empty<FitParameter>();
}

/// <summary>高级拟合选项：模型、多项式次数、平滑度、可固定/覆盖的任意参数。</summary>
public sealed class FitOptions
{
    public FitModel Model { get; set; } = FitModel.Linear;
    public int PolynomialDegree { get; set; } = 2;

    /// <summary>0..1，仅对多项式生效，值越大越平滑（抑制过拟合）。</summary>
    public double Smoothing { get; set; }

    /// <summary>每个参数位置的固定值；null 表示该参数自由拟合。</summary>
    public double?[]? FixedParams { get; set; }
}

/// <summary>
/// 对一组点做最小二乘曲线拟合。解析方法覆盖常用类型，Levenberg-Marquardt 处理
/// 非线性与任意参数固定，支持 S 型、高斯、正弦、有理函数与多项式正则化。
/// </summary>
public static class CurveFit
{
    public static int MaxPolynomialDegree => 8;

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

    public static int ParamCount(FitModel m, int degree) => m switch
    {
        FitModel.Linear => 2,
        FitModel.Polynomial => Math.Clamp(degree, 2, MaxPolynomialDegree) + 1,
        FitModel.Exp2 => 2,
        FitModel.Logarithmic => 2,
        FitModel.Power => 2,
        FitModel.Exp3 => 3,
        FitModel.Logistic => 4,
        FitModel.Gaussian => 4,
        FitModel.Sine => 4,
        FitModel.Rational => 3,
        _ => 0
    };

    public static string[] ParamNames(FitModel m, int degree)
    {
        return m switch
        {
            FitModel.Linear => new[] { "a", "b" },
            FitModel.Polynomial => Enumerable.Range(0, ParamCount(m, degree)).Select(i => $"c{i}").ToArray(),
            FitModel.Exp2 => new[] { "a", "b" },
            FitModel.Exp3 => new[] { "a", "b", "c" },
            FitModel.Logarithmic => new[] { "a", "b" },
            FitModel.Power => new[] { "a", "b" },
            FitModel.Logistic => new[] { "L", "U", "k", "m" },
            FitModel.Gaussian => new[] { "a", "b", "c", "d" },
            FitModel.Sine => new[] { "a", "b", "c", "d" },
            FitModel.Rational => new[] { "a", "b", "c" },
            _ => Array.Empty<string>()
        };
    }

    /// <summary>便捷拟合（无固定参数、无平滑）。</summary>
    public static FitResult Fit(IReadOnlyList<(double X, double Y)> pts, FitModel model, int degree = 2)
        => Fit(pts, new FitOptions { Model = model, PolynomialDegree = degree });

    /// <summary>高级拟合：支持固定/覆盖参数与多项式正则化。</summary>
    public static FitResult Fit(IReadOnlyList<(double X, double Y)> pts, FitOptions opt)
    {
        var data = pts.Where(p => double.IsFinite(p.X) && double.IsFinite(p.Y)).OrderBy(p => p.X).ToList();
        var result = new FitResult { Model = opt.Model, Label = LabelOf(opt.Model) };
        if (data.Count < 2)
        {
            result.Error = "有效拟合点不足，至少需要 2 个点";
            return result;
        }

        int degree = Math.Clamp(opt.PolynomialDegree, 2, MaxPolynomialDegree);
        int m = ParamCount(opt.Model, degree);
        var fixedParams = opt.FixedParams;
        bool hasFixed = false;
        if (fixedParams != null)
            for (int j = 0; j < m && j < fixedParams.Length; j++)
                if (fixedParams[j].HasValue) { hasFixed = true; break; }
        bool analytic = IsAnalytic(opt.Model);
        bool ridge = opt.Model == FitModel.Polynomial && opt.Smoothing > 0;

        // 快速解析路径：无固定参数、无正则化
        if (analytic && !hasFixed && !ridge)
        {
            if (TryAnalytic(data, opt.Model, degree, out var p, out var ok))
            {
                if (!ok)
                {
                    result.Error = AnalyticError(opt.Model);
                    return result;
                }
                FillResult(result, opt.Model, p, degree);
                CompleteResult(result, data, opt.Model, degree);
                return result;
            }
            result.Error = "数据点过于集中，无法拟合该线型";
            return result;
        }

        // 高级路径：非线性、固定参数或正则化
        double[] p0 = new double[m];
        if (analytic)
            TryAnalytic(data, opt.Model, degree, out p0, out _);
        else
            InitNonlinear(data, opt.Model, p0);
        if (p0.Length != m) p0 = DefaultInit(opt.Model, degree, data);
        if (p0.Any(v => !double.IsFinite(v))) p0 = DefaultInit(opt.Model, degree, data);

        if (fixedParams != null)
            for (int j = 0; j < m && j < fixedParams.Length; j++)
            {
                var fv = fixedParams[j];
                if (fv.HasValue) p0[j] = fv.Value;
            }

        bool anyActive = true;
        if (fixedParams != null && fixedParams.Length >= m)
        {
            anyActive = false;
            for (int j = 0; j < m; j++)
                if (!fixedParams[j].HasValue) { anyActive = true; break; }
        }
        if (anyActive)
        {
            var active = new bool[m];
            for (int j = 0; j < m; j++)
                active[j] = fixedParams == null || j >= fixedParams.Length || !fixedParams[j].HasValue;
            double ridgeVal = ridge ? opt.Smoothing * 1e-3 : 0;
            if (LevenbergMarquardt(data, p0, active, ModelFunc(opt.Model, degree), ridgeVal, out var p, out var err))
            {
                p0 = p;
            }
            else
            {
                result.Error = err;
                return result;
            }
        }

        FillResult(result, opt.Model, p0, degree);
        CompleteResult(result, data, opt.Model, degree);
        return result;
    }

    private static void FillResult(FitResult r, FitModel model, double[] p, int degree)
    {
        var func = ModelFunc(model, degree);
        r.Evaluate = x => func(x, p);
        r.Formula = BuildFormula(model, p, degree);
        var names = ParamNames(model, degree);
        r.Parameters = names.Select((n, i) => new FitParameter(n, p[i])).ToArray();
        r.Converged = true;
    }

    /// <summary>把参数数组包装成固定闭包，避免每次求值都解析。</summary>
    private static Func<double, double[], double> ModelFunc(FitModel model, int degree) => model switch
    {
        FitModel.Linear => (x, p) => p[0] + p[1] * x,
        FitModel.Polynomial => (x, p) => { double s = 0; int k = degree + 1; for (int i = k - 1; i >= 0; i--) s = s * x + p[i]; return s; },
        FitModel.Exp2 => (x, p) => p[0] * Math.Exp(p[1] * x),
        FitModel.Exp3 => (x, p) => p[0] + p[1] * Math.Exp(p[2] * x),
        FitModel.Logarithmic => (x, p) => p[0] + p[1] * Math.Log(x),
        FitModel.Power => (x, p) => p[0] * Math.Pow(x, p[1]),
        FitModel.Logistic => (x, p) => p[0] + (p[1] - p[0]) / (1.0 + Math.Exp(-p[2] * (x - p[3]))),
        FitModel.Gaussian => (x, p) => p[0] + p[1] * Math.Exp(-(x - p[2]) * (x - p[2]) / (2.0 * p[3] * p[3])),
        FitModel.Sine => (x, p) => p[0] + p[1] * Math.Sin(p[2] * x + p[3]),
        FitModel.Rational => (x, p) => (p[0] + p[1] * x) / (1.0 + p[2] * x),
        _ => (x, p) => 0
    };

    private static bool IsAnalytic(FitModel m)
        => m is FitModel.Linear or FitModel.Polynomial or FitModel.Exp2 or FitModel.Logarithmic or FitModel.Power;

    private static string AnalyticError(FitModel m) => m switch
    {
        FitModel.Exp2 => "双参指数要求所有 Y 均大于 0",
        FitModel.Logarithmic => "对数曲线要求所有 X 均大于 0",
        FitModel.Power => "幂函数要求所有 X 与 Y 均大于 0",
        _ => "无法拟合该线型"
    };

    private static bool TryAnalytic(List<(double X, double Y)> d, FitModel model, int degree, out double[] p, out bool ok)
    {
        ok = true;
        switch (model)
        {
            case FitModel.Linear:
            {
                var (a, b) = FitLine(d);
                p = new[] { a, b };
                return true;
            }
            case FitModel.Polynomial:
            {
                int m = degree + 1;
                if (d.Count <= degree) { p = Array.Empty<double>(); ok = false; return true; }
                var mat = new double[m, m];
                var rhs = new double[m];
                for (int i = 0; i < m; i++)
                    for (int j = 0; j < m; j++)
                    {
                        double s = 0;
                        foreach (var q in d) s += Math.Pow(q.X, i + j);
                        mat[i, j] = s;
                    }
                for (int i = 0; i < m; i++)
                {
                    double s = 0;
                    foreach (var q in d) s += q.Y * Math.Pow(q.X, i);
                    rhs[i] = s;
                }
                if (!SolveLinear(mat, rhs, out var coeff)) { p = Array.Empty<double>(); ok = false; return true; }
                p = coeff;
                return true;
            }
            case FitModel.Exp2:
            {
                if (d.Any(q => q.Y <= 0)) { p = Array.Empty<double>(); ok = false; return true; }
                var t = d.Select(q => (q.X, Math.Log(q.Y))).ToList();
                var (a, b) = FitLine(t);
                p = new[] { Math.Exp(a), b };
                return true;
            }
            case FitModel.Logarithmic:
            {
                if (d.Any(q => q.X <= 0)) { p = Array.Empty<double>(); ok = false; return true; }
                var t = d.Select(q => (Math.Log(q.X), q.Y)).ToList();
                var (a, b) = FitLine(t);
                p = new[] { a, b };
                return true;
            }
            case FitModel.Power:
            {
                if (d.Any(q => q.X <= 0 || q.Y <= 0)) { p = Array.Empty<double>(); ok = false; return true; }
                var t = d.Select(q => (Math.Log(q.X), Math.Log(q.Y))).ToList();
                var (a, b) = FitLine(t);
                p = new[] { Math.Exp(a), b };
                return true;
            }
            default:
                p = Array.Empty<double>();
                ok = false;
                return true;
        }
    }

    private static void InitNonlinear(List<(double X, double Y)> d, FitModel model, double[] p)
    {
        switch (model)
        {
            case FitModel.Exp3: InitExp3(d, p); break;
            case FitModel.Logistic: InitLogistic(d, p); break;
            case FitModel.Gaussian: InitGaussian(d, p); break;
            case FitModel.Sine: InitSine(d, p); break;
            case FitModel.Rational: InitRational(d, p); break;
            default: DefaultInit(model, 2, d).CopyTo(p, 0); break;
        }
    }

    private static double[] DefaultInit(FitModel model, int degree, List<(double X, double Y)> d)
    {
        int m = ParamCount(model, degree);
        var p = new double[m];
        double meanY = d.Average(q => q.Y);
        switch (model)
        {
            case FitModel.Linear: p[0] = meanY; break;
            case FitModel.Exp2: p[0] = Math.Max(1e-6, meanY); p[1] = 0.01; break;
            case FitModel.Exp3: p[0] = d.Min(q => q.Y); p[1] = Math.Max(1e-6, d.Max(q => q.Y) - d.Min(q => q.Y)); p[2] = 0.05; break;
            case FitModel.Logarithmic: p[0] = meanY; p[1] = 0.1; break;
            case FitModel.Power: p[0] = Math.Max(1e-6, meanY); p[1] = 1; break;
            case FitModel.Logistic: p[0] = d.Min(q => q.Y); p[1] = d.Max(q => q.Y); p[2] = 0.1; p[3] = 0.5 * (d[0].X + d[^1].X); break;
            case FitModel.Gaussian: p[0] = d.Min(q => q.Y); p[1] = Math.Max(1e-6, d.Max(q => q.Y) - d.Min(q => q.Y)); p[2] = d[0].X; p[3] = Math.Max(1e-3, (d[^1].X - d[0].X) / 5.0); break;
            case FitModel.Sine: p[0] = meanY; p[1] = Math.Max(1e-6, (d.Max(q => q.Y) - d.Min(q => q.Y)) / 2.0); p[2] = 2.0 * Math.PI / Math.Max(1e-6, d[^1].X - d[0].X); p[3] = 0; break;
            case FitModel.Rational: p[0] = meanY; p[1] = (d[^1].Y - d[0].Y) / Math.Max(1e-6, d[^1].X - d[0].X); p[2] = 0; break;
            case FitModel.Polynomial: for (int i = 0; i < m; i++) p[i] = i == 0 ? meanY : 0; break;
        }
        return p;
    }

    private static void InitExp3(List<(double X, double Y)> d, double[] p)
    {
        double minY = d.Min(q => q.Y), maxY = d.Max(q => q.Y);
        bool up = d[^1].Y >= d[0].Y;
        double range = Math.Max(1e-6, d[^1].X - d[0].X);
        double a0 = up ? minY : maxY;
        double b0 = (up ? 1.0 : -1.0) * Math.Max(1e-6, Math.Abs(maxY - minY));
        double c0 = (up ? 1.0 : -1.0) * Math.Log(2.0) / (range * 0.5);
        int i1 = Math.Max(0, d.Count / 4), i2 = Math.Min(d.Count - 1, d.Count * 3 / 4);
        double y1 = d[i1].Y - a0, y2 = d[i2].Y - a0;
        if (y1 > 1e-12 && y2 > 1e-12 && Math.Abs(d[i2].X - d[i1].X) > 1e-9)
            c0 = Math.Log(y2 / y1) / (d[i2].X - d[i1].X);
        p[0] = a0; p[1] = b0; p[2] = c0;
    }

    private static void InitLogistic(List<(double X, double Y)> d, double[] p)
    {
        double minY = d.Min(q => q.Y), maxY = d.Max(q => q.Y);
        bool up = d[^1].Y >= d[0].Y;
        double range = Math.Max(1e-6, d[^1].X - d[0].X);
        double lo = minY, hi = maxY;
        if (Math.Abs(hi - lo) < 1e-9) hi = lo + 1;
        double mid = 0.5 * (d[0].X + d[^1].X);
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
        p[0] = lo; p[1] = hi; p[2] = k0; p[3] = mid;
    }

    private static void InitGaussian(List<(double X, double Y)> d, double[] p)
    {
        double minY = d.Min(q => q.Y), maxY = d.Max(q => q.Y);
        double den = d.Sum(q => q.Y);
        double c0 = Math.Abs(den) < 1e-12 ? d[0].X : d.Sum(q => q.X * q.Y) / den;
        double d0 = Math.Max((d[^1].X - d[0].X) / 5.0, 1e-6);
        p[0] = minY; p[1] = Math.Max(1e-6, maxY - minY); p[2] = c0; p[3] = d0;
    }

    private static void InitSine(List<(double X, double Y)> d, double[] p)
    {
        double meanY = d.Average(q => q.Y);
        double range = Math.Max(1e-6, d[^1].X - d[0].X);
        p[0] = meanY;
        p[1] = Math.Max(1e-6, (d.Max(q => q.Y) - d.Min(q => q.Y)) / 2.0);
        p[2] = 2.0 * Math.PI / range;
        p[3] = 0;
    }

    private static void InitRational(List<(double X, double Y)> d, double[] p)
    {
        double meanY = d.Average(q => q.Y);
        p[0] = meanY;
        p[1] = (d[^1].Y - d[0].Y) / Math.Max(1e-6, d[^1].X - d[0].X);
        p[2] = 0;
    }

    private static string BuildFormula(FitModel model, double[] p, int degree) => model switch
    {
        FitModel.Linear => $"y = {F(p[0])} {(p[1] >= 0 ? "+" : "-")} {F(Math.Abs(p[1]))}x",
        FitModel.Polynomial => BuildPolyFormula(p),
        FitModel.Exp2 => $"y = {F(p[0])} * e^({F(p[1])}x)",
        FitModel.Exp3 => $"y = {F(p[0])} + {F(p[1])}*e^({F(p[2])}x)",
        FitModel.Logarithmic => $"y = {F(p[0])} {(p[1] >= 0 ? "+" : "-")} {F(Math.Abs(p[1]))}*ln(x)",
        FitModel.Power => $"y = {F(p[0])} * x^{F(p[1])}",
        FitModel.Logistic => $"y = {F(p[0])} + ({F(p[1] - p[0])})/(1+e^(-{F(p[2])}(x-{F(p[3])})))",
        FitModel.Gaussian => $"y = {F(p[0])} + {F(p[1])}*e^(-((x-{F(p[2])})^2)/{F(2.0 * p[3] * p[3])})",
        FitModel.Sine => $"y = {F(p[0])} + {F(p[1])}*sin({F(p[2])}x+{F(p[3])})",
        FitModel.Rational => $"y = ({F(p[0])}+{F(p[1])}x)/(1+{F(p[2])}x)",
        _ => ""
    };

    private static void CompleteResult(FitResult r, List<(double X, double Y)> d, FitModel model, int degree)
    {
        if (r.Evaluate == null) return;
        var (r2, rmse) = Goodness(d, r.Evaluate);
        r.R2 = r2;
        r.RMSE = rmse;
    }

    /// <summary>Levenberg-Marquardt：数值雅可比 + 阻尼高斯-牛顿，支持固定参数与正则化。</summary>
    private static bool LevenbergMarquardt(
        List<(double X, double Y)> d,
        double[] p0,
        bool[] active,
        Func<double, double[], double> model,
        double ridge,
        out double[] p,
        out string error)
    {
        error = "";
        int n = d.Count, m = p0.Length;
        var act = new List<int>();
        for (int j = 0; j < m; j++) if (active[j]) act.Add(j);
        int ma = act.Count;
        if (n < ma)
        {
            error = $"该线型至少需要 {Math.Max(2, ma)} 个点";
            p = (double[])p0.Clone();
            return false;
        }
        p = (double[])p0.Clone();
        double lambda = 1e-3;
        double current = Sse(d, p, model, ridge);
        if (!double.IsFinite(current)) current = double.MaxValue / 2;

        for (int iter = 0; iter < 150; iter++)
        {
            if (lambda > 1e12)
            {
                error = "未收敛（阻尼过大），请换线型或检查数据";
                return false;
            }
            var jac = new double[n, ma];
            for (int i = 0; i < n; i++)
            {
                double xi = d[i].X;
                for (int k = 0; k < ma; k++)
                {
                    int j = act[k];
                    double h = 1e-6 * Math.Max(1.0, Math.Abs(p[j]));
                    var qp = (double[])p.Clone(); qp[j] += h;
                    var qm = (double[])p.Clone(); qm[j] -= h;
                    double fp = model(xi, qp), fm = model(xi, qm);
                    if (double.IsFinite(fp) && double.IsFinite(fm))
                        jac[i, k] = (fp - fm) / (2.0 * h);
                    else
                        jac[i, k] = 0;
                }
            }

            var a = new double[ma, ma];
            var g = new double[ma];
            for (int i = 0; i < n; i++)
            {
                double ri = d[i].Y - model(d[i].X, p);
                if (!double.IsFinite(ri)) ri = 0;
                for (int k = 0; k < ma; k++)
                {
                    g[k] += jac[i, k] * ri;
                    for (int l = k; l < ma; l++)
                        a[k, l] += jac[i, k] * jac[i, l];
                }
            }
            for (int k = 0; k < ma; k++)
            {
                int j = act[k];
                if (j >= 1 && ridge > 0)
                {
                    double c = 2.0 * ridge;
                    a[k, k] += c;
                    g[k] += c * p[j];
                }
                a[k, k] = a[k, k] * (1.0 + lambda) + 1e-12;
                for (int l = k + 1; l < ma; l++) a[l, k] = a[k, l];
            }

            if (!SolveLinear(a, g, out var delta))
            {
                lambda *= 10.0;
                continue;
            }
            double deltaMax = delta.Length == 0 ? 0 : delta.Max(Math.Abs);
            if (deltaMax < 1e-12) break; // 已到最优或梯度消失
            var candidate = (double[])p.Clone();
            for (int k = 0; k < ma; k++) candidate[act[k]] += delta[k];
            double next = Sse(d, candidate, model, ridge);
            if (double.IsFinite(next) && next < current)
            {
                double improvement = Math.Abs(current - next) / Math.Max(1.0, current);
                p = candidate;
                current = next;
                lambda = Math.Max(lambda / 3.0, 1e-12);
                if (improvement < 1e-11) break;
            }
            else
            {
                lambda *= 5.0;
            }
        }
        return true;
    }

    private static double Sse(List<(double X, double Y)> d, double[] p, Func<double, double[], double> model, double ridge)
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
        if (ridge > 0)
            for (int j = 1; j < p.Length; j++) s += ridge * p[j] * p[j];
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
