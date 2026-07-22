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
    private const int MaxBulletsPerTower = 5;
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(10) };
    private readonly TeamConfig _teamConfig;

    public OllamaService(TeamConfig teamConfig) => _teamConfig = teamConfig;

    public event Action<string>? StatusUpdate;

    // Per-person summarization (the expensive Ollama step): one filtered bullet list per reporter.
    private async Task<List<(string Name, string Summary)>> SummarizePeopleAsync(string weekLabel, List<EmailItem> emails, CancellationToken ct)
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

            // Deterministic short-circuit: a brief "no work this period" note (PTO/OOO) becomes a
            // single accurate bullet, bypassing the model so it can't pad it with fabricated activity.
            var combinedBody = string.Join("\n", group.Select(e => e.Body));
            if (IsNoWorkUpdate(combinedBody))
            {
                personSummaries.Add((group.Key, "- " + NoWorkBullet(combinedBody)));
                continue;
            }

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

                Extract only the activities, tasks, blockers, and next steps this person actually mentioned.
                Output as concise bullet points starting with -. No intro text. Bullets only.
                Create a separate bullet for EACH distinct activity or update that is actually stated. Do NOT
                merge separate points together and do NOT skip any. Produce up to 8 bullets. Do NOT invent,
                pad, or add filler, and do NOT output "no updates" / "nothing to report" style bullets.
                Keep the bullets in the SAME top-to-bottom order they appear in the update — do not reorder
                them. The points nearer the top of the update are the most important, so list them first.
                Do NOT create a bullet for server, VM, instance, or database counts (for example
                "22 instances on RISE with 110 servers", "Total VMs: 126", "RHEL - 31") — those totals are
                reported in a separate section, so leave them out here.
                If the update is brief or just says the person had no work this period (for example, on PTO),
                output a single short bullet stating exactly that and nothing more.
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
                .Where(l => !IsFillerBullet(l))
                .Where(l => !IsCountLine(l))
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

        return personSummaries;
    }

    // Step 1: candidate data points per tower (top-first order), for the user to pick from.
    public async Task<Dictionary<string, List<string>>> ExtractDataPointsAsync(string weekLabel, List<EmailItem> emails, CancellationToken ct)
    {
        var personSummaries = await SummarizePeopleAsync(weekLabel, emails, ct);
        var byTower = personSummaries
            .GroupBy(p => _teamConfig.GetTeam(p.Name))
            .ToDictionary(g => g.Key, g => g.ToList());

        var result = new Dictionary<string, List<string>>();
        foreach (var tower in TeamConfig.TowerNames)
        {
            if (!byTower.TryGetValue(tower, out var members) || members.Count == 0) continue;
            var candidates = TopBulletsPerTower(members, int.MaxValue); // all candidates, top-first
            if (candidates.Count > 0) result[tower] = candidates;
        }
        return result;
    }

    // Step 2: assemble the final report from the data points the user selected per tower.
    public async Task<string> BuildReportAsync(string weekLabel, Dictionary<string, List<string>> selectedByTower, CancellationToken ct)
    {
        var teamSection = new StringBuilder();
        var towerSection = new StringBuilder();
        foreach (var tower in TeamConfig.TowerNames)
        {
            if (!selectedByTower.TryGetValue(tower, out var bullets) || bullets.Count == 0) continue;
            teamSection.AppendLine($"**{tower}**");
            towerSection.AppendLine($"**{tower}**");
            foreach (var b in bullets)
            {
                teamSection.AppendLine($"- {b}");
                towerSection.AppendLine($"- {b}");
            }
            teamSection.AppendLine();
            towerSection.AppendLine();
        }

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
    /// Pulls the server / database counts the towers report each week (SAP RISE/Xeta instances
    /// and servers, SQL databases, Cloud Total VMs) for the lower-right template table.
    /// All parsing is deterministic regex over the tower emails — no model, so the numbers are exact.
    /// </summary>
    public Task<TowerMetrics> ExtractMetricsAsync(List<EmailItem> emails, CancellationToken ct)
    {
        var metrics = new TowerMetrics();
        var byTower = emails
            .Where(e => !e.From.Contains("Coopersmith", StringComparison.OrdinalIgnoreCase))
            .GroupBy(e => _teamConfig.GetTeam(e.From))
            .ToDictionary(g => g.Key, g => g.ToList());

        if (byTower.TryGetValue("SAP", out var sapEmails))
        {
            StatusUpdate?.Invoke("Reading SAP instance/server counts...");
            var text = CombineBodies(sapEmails);
            var rise = SapRiseRegex.Match(text);
            if (rise.Success)
            {
                metrics.SapRiseInstances = ParseInt(rise.Groups[1]);
                metrics.SapRiseServers = ParseInt(rise.Groups[2]);
            }
            var xeta = SapXetaRegex.Match(text);
            if (xeta.Success)
            {
                metrics.SapXetaInstances = ParseInt(xeta.Groups[1]);
                metrics.SapXetaServers = ParseInt(xeta.Groups[2]);
            }
        }

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

        return Task.FromResult(metrics);
    }

    // Round-robin across a tower's members (one bullet each per pass) up to max, so no tower
    // ever shows more than `max` bullets and multiple reporters are represented fairly.
    private static List<string> TopBulletsPerTower(List<(string Name, string Summary)> members, int max)
    {
        var perMember = members
            .Select(m => m.Summary.Split('\n')
                .Select(l => l.TrimStart().TrimStart('-', '•').Trim())
                .Where(l => l.Length > 0)
                .ToList())
            .Where(l => l.Count > 0)
            .ToList();

        var result = new List<string>();
        for (int round = 0; result.Count < max; round++)
        {
            bool added = false;
            foreach (var bl in perMember)
            {
                if (round < bl.Count)
                {
                    result.Add(bl[round]);
                    added = true;
                    if (result.Count >= max) break;
                }
            }
            if (!added) break;
        }
        return result;
    }

    private static string CombineBodies(List<EmailItem> emails)
    {
        var sb = new StringBuilder();
        foreach (var e in emails)
        {
            var excerpt = e.Body.Length > 6000 ? e.Body[..6000] : e.Body;
            sb.AppendLine($"From: {e.From}");
            sb.AppendLine($"Subject: {e.Subject}");
            sb.AppendLine(excerpt);
            sb.AppendLine();
        }
        return sb.ToString();
    }

    // SAP: "22 instances on SAP RISE with 110 servers" and "4 instances on Xeta with 19 servers"
    private static readonly Regex SapRiseRegex = new(@"(\d[\d,]*)\s+instances?\s+on\s+(?:sap\s+)?rise\s+with\s+(\d[\d,]*)\s+servers?", RegexOptions.IgnoreCase);
    private static readonly Regex SapXetaRegex = new(@"(\d[\d,]*)\s+instances?\s+on\s+xeta\s+with\s+(\d[\d,]*)\s+servers?", RegexOptions.IgnoreCase);

    private static int? ParseInt(Group g) => int.TryParse(g.Value.Replace(",", ""), out var n) ? n : null;

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

    // A brief note whose only substance is "I was out" (PTO / OOO / on leave / no updates).
    // Length-gated so a real update that merely mentions upcoming PTO isn't caught.
    private static bool IsNoWorkUpdate(string body)
    {
        var b = Regex.Replace(body ?? "", @"\s+", " ").Trim();
        if (b.Length == 0 || b.Length > 200) return false;
        return Regex.IsMatch(b, @"\b(pto|out\s+of\s+office|o\.?o\.?o\.?|on\s+leave|on\s+vacation|was\s+off|no\s+updates?|nothing\s+to\s+report)\b",
            RegexOptions.IgnoreCase);
    }

    // Turns that brief note into one clean bullet (greeting stripped).
    private static string NoWorkBullet(string body)
    {
        var b = Regex.Replace(body ?? "", @"\s+", " ").Trim();
        b = Regex.Replace(b, @"^(hi|hello|hey|good\s+morning|good\s+afternoon|greetings)\b[^,.\n]*[,:]?\s*", "",
            RegexOptions.IgnoreCase).Trim().TrimEnd('.', ' ');
        if (b.Length == 0) return "On PTO";
        return b.Length > 120 ? b.Substring(0, 120).Trim() : b;
    }

    // Count / inventory lines that belong in the Managed Services Tasks pane, not the accomplishments
    // narrative: SAP RISE/Xeta instances & servers, DB counts, Total VMs, "N servers/VMs/instances",
    // "Servers: N", and bare "RHEL - 31" tallies. Kept in sync with PowerPointService.MetricLinePatterns.
    private static readonly Regex[] CountLinePatterns =
    {
        new(@"\d[\d,]*\s+instances?\s+on\s+(?:sap\s+)?(?:rise|xeta)\b", RegexOptions.IgnoreCase),
        new(@"\d[\d,]*\s+(?:sql\s+)?(?:databases?|dbs?)\b", RegexOptions.IgnoreCase),
        new(@"total\s+vm['’]?s?\b", RegexOptions.IgnoreCase),
        new(@"\b\d[\d,]*\s+(?:[A-Za-z]+\s+)?(?:servers?|vms?|virtual\s+machines?|instances?|nodes?|hosts?|machines?)\b", RegexOptions.IgnoreCase),
        new(@"\b(?:servers?|vms?|instances?|databases?|dbs?|nodes?|hosts?|machines?)\s*[-:]\s*\d", RegexOptions.IgnoreCase),
        new(@"^[A-Za-z][A-Za-z ._/]*\s*[-:]\s*\d[\d,]*\s*$", RegexOptions.IgnoreCase),
    };

    private static bool IsCountLine(string line)
    {
        var s = line.TrimStart(' ', '\t', '-', '•', '*').Trim();
        return CountLinePatterns.Any(rx => rx.IsMatch(s));
    }

    // True for "absence of content" filler the model sometimes emits when an update is trivial
    // (e.g. "No updates mentioned", "No specific tasks or activities mentioned", "Nothing to report").
    // A real bullet like "On PTO last week" is kept.
    private static bool IsFillerBullet(string line)
    {
        var s = line.TrimStart(' ', '\t', '-', '•', '*').Trim().TrimEnd('.').ToLowerInvariant();
        if (s.Length == 0) return true;
        if (s is "n/a" or "na" or "none" or "nothing to report" or "not applicable" or "no updates" or "no update")
            return true;
        if (s.StartsWith("no ") &&
            Regex.IsMatch(s, @"\b(mentioned|reported|provided|noted|stated|listed|specified|available)\b"))
            return true;
        if (Regex.IsMatch(s, @"^no\s+(specific\s+)?(updates?|tasks?|activities|activity|items?|details?)\s*$"))
            return true;
        return false;
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
