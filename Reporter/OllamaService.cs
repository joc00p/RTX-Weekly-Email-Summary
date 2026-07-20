using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Threading;
using System.Threading.Tasks;

namespace Reporter;

public class OllamaService
{
    private const string Url = "http://localhost:11434/api/generate";
    private const string Model = "llama3.2";
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(10) };
    private readonly TeamConfig _teamConfig;

    public OllamaService(TeamConfig teamConfig) => _teamConfig = teamConfig;

    public event Action<string>? StatusUpdate;

    public async Task<string> SummarizeWeekAsync(string weekLabel, List<EmailItem> emails, CancellationToken ct)
    {
        // Group emails by sender, exclude Joel Coopersmith
        var bySender = emails
            .Where(e => !e.From.Contains("Coopersmith", StringComparison.OrdinalIgnoreCase))
            .GroupBy(e => e.From)
            .OrderBy(g => g.Key)
            .ToList();

        // Step 1: Summarize each person individually
        var personSummaries = new List<(string Name, string Summary)>();
        foreach (var group in bySender)
        {
            ct.ThrowIfCancellationRequested();
            StatusUpdate?.Invoke($"Summarizing {group.Key} ({personSummaries.Count + 1}/{bySender.Count})...");

            var blocks = new StringBuilder();
            foreach (var e in group)
            {
                var excerpt = e.Body.Length > 6000 ? e.Body[..6000] : e.Body;
                blocks.AppendLine($"Date: {e.Received}");
                blocks.AppendLine($"Subject: {e.Subject}");
                blocks.AppendLine(excerpt);
                blocks.AppendLine();
            }

            var prompt = $"""
                You are summarizing work updates from one team member for a status report.
                Person: {group.Key}
                Period: {weekLabel}

                Their submitted updates:
                {blocks}

                Extract the activities, tasks, blockers, and next steps this person mentioned.
                Output as concise bullet points starting with -. No invented content. No intro text. Bullets only.
                Capture each distinct point as its own bullet — do not merge separate points together.
                Produce up to 6 bullets (aim for 5 to 6 when the update has that many distinct points), and at least one.
                If details are sparse, summarize what little was provided.
                Do NOT use the phrase "punch list" or "punch lists" anywhere in your response.
                """;

            var summary = await CallOllama(prompt, ct);
            var trimmed = summary.Trim();
            // Always include this person — fall back to a raw excerpt if Ollama produced nothing usable
            if (string.IsNullOrWhiteSpace(trimmed) || trimmed.StartsWith("(no updates)", StringComparison.OrdinalIgnoreCase))
            {
                var fallbackLine = group.First().Subject;
                trimmed = $"- {(string.IsNullOrWhiteSpace(fallbackLine) ? "Updates submitted" : fallbackLine)}";
            }
            // Hard cap: keep at most 6 bullet lines; drop any mentioning Coopersmith
            var bulletLines = trimmed.Split('\n')
                .Where(l => l.TrimStart().StartsWith('-') || l.TrimStart().StartsWith('•'))
                .Where(l => !l.Contains("Coopersmith", StringComparison.OrdinalIgnoreCase))
                .Where(l => !l.Contains("punch list", StringComparison.OrdinalIgnoreCase))
                .Take(6)
                .ToList();
            if (bulletLines.Count == 0)
            {
                var fallbackLine = group.First().Subject;
                trimmed = $"- {(string.IsNullOrWhiteSpace(fallbackLine) ? "Updates submitted" : fallbackLine)}";
            }
            else
                trimmed = string.Join("\n", bulletLines);
            personSummaries.Add((group.Key, trimmed));
        }

        // Step 2: Build the full report from individual summaries
        StatusUpdate?.Invoke("Assembling report...");
        var teamSection = new StringBuilder();
        foreach (var (name, summary) in personSummaries)
        {
            teamSection.AppendLine($"**{name}**");
            teamSection.AppendLine(summary);
            teamSection.AppendLine();
        }

        // Step 3: Build a tower-grouped section for the exec prompt (no individual names)
        var towerOrder = TeamConfig.TowerNames.Append("Other").ToArray();
        var byTower = personSummaries
            .GroupBy(p => _teamConfig.GetTeam(p.Name))
            .ToDictionary(g => g.Key, g => g.ToList());

        var towerSection = new StringBuilder();
        foreach (var tower in towerOrder)
        {
            if (!byTower.TryGetValue(tower, out var members) || members.Count == 0) continue;
            towerSection.AppendLine($"**{tower}**");
            foreach (var (_, summary) in members)
                towerSection.AppendLine(summary);
            towerSection.AppendLine();
        }

        // Step 4: Generate summary + executive summary from tower-grouped sections
        StatusUpdate?.Invoke("Writing executive summary...");
        var execPrompt = $"""
            You are writing the final sections of a team status report for the period: {weekLabel}

            Tower updates (do NOT mention individual names — refer only to the tower names):
            {towerSection}

            Do NOT use the phrase "punch list" or "punch lists" anywhere in your response.
            Do NOT mention any individual person's name. Refer only to the tower (SAP, Cloud, DBA SQL, ITIL Svc Mgmt, etc.).
            Do NOT include any mention of risks, issues, blockers, or problems — omit them entirely.
            Write ONLY these two sections exactly as formatted:

            ### Summary
            - [overall team progress]
            - [key accomplishments]
            - [items pending]

            ### Executive Summary
            [3-5 sentences of high-level professional prose for senior leadership. Write about what was accomplished — the actual work, projects, and deliverables — not about which team or tower did it. Do not use tower names as sentence subjects or anchors. Do not list, enumerate, or organize the summary around towers. Do not mention risks, issues, or blockers. Do not name individuals.]
            """;

        var execSummary = StripRiskLines(await CallOllama(execPrompt, ct));

        // Assemble final report
        var report = new StringBuilder();
        report.AppendLine($"## STATUS REPORT — {weekLabel}");
        report.AppendLine();
        report.AppendLine("### Team Updates");
        report.AppendLine();
        report.Append(teamSection);
        report.AppendLine("---");
        report.AppendLine();
        report.AppendLine(execSummary.Trim());

        return CleanBulletSpacing(RemoveBannedPhrases(report.ToString()));
    }

    /// <summary>
    /// Pulls the server / database counts the towers report each week (SAP instances &amp;
    /// RISE/XETA servers, SQL databases, Cloud servers) for the lower-right template table.
    /// One focused extraction per tower keeps the local model's context tight.
    /// </summary>
    public async Task<TowerMetrics> ExtractMetricsAsync(List<EmailItem> emails, CancellationToken ct)
    {
        var metrics = new TowerMetrics();
        var byTower = emails
            .Where(e => !e.From.Contains("Coopersmith", StringComparison.OrdinalIgnoreCase))
            .GroupBy(e => _teamConfig.GetTeam(e.From))
            .ToDictionary(g => g.Key, g => g.ToList());

        if (byTower.TryGetValue("SAP", out var sapEmails))
        {
            ct.ThrowIfCancellationRequested();
            StatusUpdate?.Invoke("Extracting SAP server counts...");
            var d = ParseNumbers(await CallOllama(SapMetricsPrompt(CombineBodies(sapEmails)), ct));
            metrics.SapInstances    = Pick(d, "instances");
            metrics.SapRiseServers  = Pick(d, "rise_servers");
            metrics.SapRiseLiveApps = Pick(d, "rise_live_apps");
            metrics.SapXetaServers  = Pick(d, "xeta_servers");
            metrics.SapXetaLiveApps = Pick(d, "xeta_live_apps");
        }

        // DBA and Cloud are summed deterministically (no model): add up every labeled
        // part found in that tower's emails. Repeatable and not subject to model guessing.
        if (byTower.TryGetValue("DBA SQL", out var dbaEmails))
        {
            StatusUpdate?.Invoke("Summing SQL database counts...");
            (metrics.SqlDatabases, metrics.SqlBreakdown) = SumMatches(CombineBodies(dbaEmails), DatabaseCountRegex);
        }

        if (byTower.TryGetValue("CLOUD", out var cloudEmails))
        {
            StatusUpdate?.Invoke("Reading Cloud Total VMs...");
            metrics.CloudServers = ExtractTotalVms(CombineBodies(cloudEmails));
        }

        return metrics;

        static int? Pick(Dictionary<string, int?> d, string key) =>
            d.TryGetValue(key, out var v) ? v : null;
    }

    private static string CombineBodies(List<EmailItem> emails)
    {
        var sb = new StringBuilder();
        foreach (var e in emails)
        {
            var excerpt = e.Body.Length > 2000 ? e.Body[..2000] : e.Body;
            sb.AppendLine($"From: {e.From}");
            sb.AppendLine($"Subject: {e.Subject}");
            sb.AppendLine(excerpt);
            sb.AppendLine();
        }
        return sb.ToString();
    }

    private static string SapMetricsPrompt(string text) => $$"""
        Extract exact numbers from this SAP team status update.
        Return ONLY a JSON object, nothing else, in exactly this shape:
        {"instances": 0, "rise_servers": 0, "rise_live_apps": 0, "xeta_servers": 0, "xeta_live_apps": 0}
        Use null for any value that is not explicitly stated. Do not guess or calculate.

        SAP update:
        {{text}}
        """;

    // Cloud total = the "Total VMs" figure the Cloud tower reports (number after or before the label).
    private static readonly Regex TotalVmsAfterRegex = new(@"total\s+vm['’]?s?\b\s*[:=\-]?\s*(\d[\d,]*)", RegexOptions.IgnoreCase);
    private static readonly Regex TotalVmsBeforeRegex = new(@"(\d[\d,]*)\s+total\s+vm['’]?s?\b", RegexOptions.IgnoreCase);
    // "<number> databases", "<number> SQL databases", or "<number> DBs" — summed for the DBA total.
    private static readonly Regex DatabaseCountRegex = new(@"(\d[\d,]*)\s+(?:sql\s+)?(?:databases?|dbs?)\b", RegexOptions.IgnoreCase);

    // Sums every number that precedes the target noun. Returns the total and a "a + b + c"
    // breakdown for display, or (null, null) when nothing matched (caller keeps template value).
    private static (int? Total, string? Breakdown) SumMatches(string text, Regex rx)
    {
        var parts = new List<int>();
        foreach (Match m in rx.Matches(text))
            if (int.TryParse(m.Groups[1].Value.Replace(",", ""), out var n))
                parts.Add(n);

        if (parts.Count == 0) return (null, null);
        return (parts.Sum(), string.Join(" + ", parts));
    }

    // Reads the single "Total VMs" figure from the Cloud tower email (label may lead or trail the number).
    private static int? ExtractTotalVms(string text)
    {
        var m = TotalVmsAfterRegex.Match(text);
        if (!m.Success) m = TotalVmsBeforeRegex.Match(text);
        return m.Success && int.TryParse(m.Groups[1].Value.Replace(",", ""), out var n) ? n : null;
    }

    // Pulls integer values out of the model's JSON reply, tolerating extra prose around it.
    private static Dictionary<string, int?> ParseNumbers(string response)
    {
        var result = new Dictionary<string, int?>(StringComparer.OrdinalIgnoreCase);
        if (string.IsNullOrWhiteSpace(response)) return result;

        int start = response.IndexOf('{');
        int end = response.LastIndexOf('}');
        if (start < 0 || end <= start) return result;

        try
        {
            using var doc = JsonDocument.Parse(response.Substring(start, end - start + 1));
            foreach (var prop in doc.RootElement.EnumerateObject())
            {
                result[prop.Name] = prop.Value.ValueKind switch
                {
                    JsonValueKind.Number => prop.Value.TryGetInt32(out var n) ? n : null,
                    JsonValueKind.String => TryExtractInt(prop.Value.GetString()),
                    _ => null,
                };
            }
        }
        catch { }
        return result;
    }

    private static int? TryExtractInt(string? s)
    {
        if (string.IsNullOrWhiteSpace(s)) return null;
        var m = Regex.Match(s, @"\d+");
        return m.Success && int.TryParse(m.Value, out var n) ? n : null;
    }

    // Matches a risk *label* line: an optional bullet, a risk keyword, then a colon/dash.
    // e.g. "Risks/issues:", "- Risk: foo", "Blockers -", "Open issues:"
    private static readonly Regex RiskLabelRegex = new(
        @"^\s*(?:[-•*]\s*)?(?:risks?|blockers?|issues?|open\s+issues?|outstanding\s+issues?|concerns?|risks?\s*/\s*issues?)\s*[:\-–—]",
        RegexOptions.IgnoreCase);

    // Matches a bare risk *heading* line with no trailing detail. e.g. "Risks", "Risks/Issues"
    private static readonly Regex RiskHeadingRegex = new(
        @"^\s*(?:[-•*]\s*)?(?:risks?|blockers?|risks?\s*/\s*issues?)\s*$",
        RegexOptions.IgnoreCase);

    // Removes hallucinated risk sections without touching prose that merely mentions
    // an "issue" (e.g. "resolved the DNS issue", "certificate was issued").
    private static string StripRiskLines(string text)
    {
        var lines = text.Split('\n');
        var kept = new List<string>();
        bool inRiskBlock = false;

        foreach (var line in lines)
        {
            // A risk label/heading opens a block we drop (the label plus any bullets under it)
            if (RiskLabelRegex.IsMatch(line) || RiskHeadingRegex.IsMatch(line))
            {
                inRiskBlock = true;
                continue;
            }

            if (inRiskBlock)
            {
                var t = line.TrimStart();
                bool isIndented = line.Length > 0 && t.Length < line.Length;
                // Stay in block for: bullets, numbered items, bold markers, blank lines, indented continuations
                if (string.IsNullOrWhiteSpace(t) ||
                    t.StartsWith('-') || t.StartsWith('•') || t.StartsWith('*') ||
                    t.StartsWith("**") ||
                    Regex.IsMatch(t, @"^\d+\.") ||
                    isIndented)
                    continue;
                inRiskBlock = false; // block ended — fall through and keep this line
            }

            kept.Add(line);
        }

        return string.Join("\n", kept);
    }

    private static string RemoveBannedPhrases(string text)
    {
        var lines = text.Split('\n');
        var kept = lines.Where(l =>
            !l.Contains("punch list", StringComparison.OrdinalIgnoreCase));
        return string.Join("\n", kept);
    }

    private static string CleanBulletSpacing(string text)
    {
        var lines = text.Split('\n');
        var result = new List<string>();

        for (int i = 0; i < lines.Length; i++)
        {
            var line = lines[i].TrimEnd();

            // Skip blank lines that fall between two bullet lines
            if (string.IsNullOrWhiteSpace(line))
            {
                int next = i + 1;
                while (next < lines.Length && string.IsNullOrWhiteSpace(lines[next]))
                    next++;

                bool prevIsBullet = result.Count > 0 && lines[i - 1].TrimStart().StartsWith('-');
                bool nextIsBullet = next < lines.Length && lines[next].TrimStart().StartsWith('-');

                if (prevIsBullet && nextIsBullet)
                {
                    i = next - 1; // skip all blanks, next iteration picks up the bullet
                    continue;
                }

                // Collapse multiple blank lines into one
                result.Add("");
                i = next - 1;
                continue;
            }

            result.Add(line);
        }

        return string.Join("\n", result).Trim();
    }

    private static async Task<string> CallOllama(string prompt, CancellationToken ct)
    {
        var payload = JsonSerializer.Serialize(new
        {
            model = Model,
            prompt,
            stream = false,
            options = new { num_ctx = 16384 }
        });
        var content = new StringContent(payload, Encoding.UTF8, "application/json");
        var response = await Http.PostAsync(Url, content, ct);
        response.EnsureSuccessStatusCode();
        var json = await response.Content.ReadAsStringAsync(ct);
        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.GetProperty("response").GetString() ?? "";
    }
}
