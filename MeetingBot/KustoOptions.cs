namespace MeetingBot;

/// <summary>Configuration for querying Azure Data Explorer telemetry.</summary>
public sealed class KustoOptions
{
    /// <summary>The Azure Data Explorer cluster URI.</summary>
    public string ClusterUri { get; set; } = string.Empty;

    /// <summary>The database containing meeting telemetry.</summary>
    public string Database { get; set; } = string.Empty;
}
