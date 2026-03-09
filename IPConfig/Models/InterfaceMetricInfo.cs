namespace IPConfig.Models;

public readonly record struct InterfaceMetricInfo(int? Metric, bool? AutomaticMetric)
{
    public string ToDisplayString()
    {
        if (AutomaticMetric is null)
        {
            return "Not available";
        }

        return AutomaticMetric.Value
            ? "Automatic"
            : Metric is int metric
                ? $"Manual ({metric})"
                : "Manual";
    }
}
