namespace GameCurve.Services;

/// <summary>一组值的统计摘要。</summary>
public readonly record struct StatSummary(int Count, double Min, double Max, double Avg, double Median, double Sum, double Std)
{
    public string AsText() =>
        $"数量:{Count}  最小:{Min:0.###}  最大:{Max:0.###}\n" +
        $"平均:{Avg:0.###}  中位:{Median:0.###}\n" +
        $"总和:{Sum:0.###}  标准差:{Std:0.###}";
}

public static class Statistics
{
    public static StatSummary Summarize(IEnumerable<double> values)
    {
        var arr = values.Where(d => !double.IsNaN(d) && !double.IsInfinity(d)).ToList();
        if (arr.Count == 0) return new StatSummary(0, 0, 0, 0, 0, 0, 0);
        arr.Sort();
        double min = arr[0], max = arr[^1], sum = arr.Sum();
        double avg = sum / arr.Count;
        int n = arr.Count;
        double median = n % 2 == 1 ? arr[n / 2] : (arr[n / 2 - 1] + arr[n / 2]) / 2.0;
        double variance = arr.Sum(v => (v - avg) * (v - avg)) / n;
        double std = Math.Sqrt(variance);
        return new StatSummary(n, min, max, avg, median, sum, std);
    }
}
