using GameCurve.Models;

namespace GameCurve.Services;

/// <summary>
/// 曲线相关的数学计算：Catmull-Rom 样条插值、平滑、归一化、钳制。
/// </summary>
public static class CurveMath
{
    /// <summary>
    /// 用 Catmull-Rom 样条对点列插值，返回用于描边的平滑折线（顺序按 X 升序）。
    /// </summary>
    public static List<(double X, double Y)> CatmullRom(IReadOnlyList<CurvePoint> points, int segmentsPerSegment = 24)
    {
        var ordered = points.OrderBy(p => p.X).ToList();
        if (ordered.Count == 0) return new List<(double, double)>();
        if (ordered.Count == 1) return new List<(double, double)> { (ordered[0].X, ordered[0].Y) };
        if (ordered.Count == 2) return new List<(double, double)>
        {
            (ordered[0].X, ordered[0].Y),
            (ordered[1].X, ordered[1].Y)
        };

        var result = new List<(double X, double Y)>();
        result.Add((ordered[0].X, ordered[0].Y));

        for (int i = 0; i < ordered.Count - 1; i++)
        {
            var p0 = ordered[Math.Max(0, i - 1)];
            var p1 = ordered[i];
            var p2 = ordered[i + 1];
            var p3 = ordered[Math.Min(ordered.Count - 1, i + 2)];

            for (int s = 1; s <= segmentsPerSegment; s++)
            {
                double t = (double)s / segmentsPerSegment;
                double t2 = t * t, t3 = t2 * t;
                double x = 0.5 * ((2 * p1.X) + (-p0.X + p2.X) * t + (2 * p0.X - 5 * p1.X + 4 * p2.X - p3.X) * t2 + (-p0.X + 3 * p1.X - 3 * p2.X + p3.X) * t3);
                double y = 0.5 * ((2 * p1.Y) + (-p0.Y + p2.Y) * t + (2 * p0.Y - 5 * p1.Y + 4 * p2.Y - p3.Y) * t2 + (-p0.Y + 3 * p1.Y - 3 * p2.Y + p3.Y) * t3);
                result.Add((x, y));
            }
        }
        return result;
    }

    /// <summary>对一组数值做窗口均值平滑（边界不回绕）。</summary>
    public static double[] MovingAverage(IReadOnlyList<double> values, int window)
    {
        int n = values.Count;
        if (n == 0) return Array.Empty<double>();
        window = Math.Clamp(window, 1, n);
        var result = new double[n];
        int half = window / 2;
        for (int i = 0; i < n; i++)
        {
            int start = Math.Max(0, i - half);
            int end = Math.Min(n - 1, i + half);
            double sum = 0; int cnt = 0;
            for (int j = start; j <= end; j++) { sum += values[j]; cnt++; }
            result[i] = sum / cnt;
        }
        return result;
    }

    public static double Clamp(double v, double min, double max)
        => Math.Max(min, Math.Min(max, v));
}
