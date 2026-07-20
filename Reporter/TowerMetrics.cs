namespace Reporter;

/// <summary>
/// Server / database counts pulled from the weekly tower emails and rendered into the
/// "Managed Services Tasks" table (lower-right) of the export template.
/// A null value means the number wasn't found in this week's emails — in that case the
/// template's existing number is left untouched rather than blanked or guessed.
/// </summary>
public class TowerMetrics
{
    // SAP — instance and server counts for RISE and Xeta
    public int? SapRiseInstances { get; set; }
    public int? SapRiseServers { get; set; }
    public int? SapXetaInstances { get; set; }
    public int? SapXetaServers { get; set; }

    // DBA SQL — number of databases (summed from the per-part figures in the emails)
    public int? SqlDatabases { get; set; }
    public string? SqlBreakdown { get; set; }   // e.g. "52 + 48 + 97"

    // CLOUD — number of servers summed across every environment listed in the emails
    public int? CloudServers { get; set; }
    public string? CloudBreakdown { get; set; } // e.g. "40 + 33 + 40"

    public bool AnyFound =>
        SapRiseInstances.HasValue || SapRiseServers.HasValue ||
        SapXetaInstances.HasValue || SapXetaServers.HasValue ||
        SqlDatabases.HasValue || CloudServers.HasValue;
}
