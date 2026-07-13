using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace RTXReporter;

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
                var excerpt = e.Body.Length > 1500 ? e.Body[..1500] : e.Body;
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

                Extract the most important activities, tasks, blockers, and next steps this person mentioned.
                Output as concise bullet points starting with -. No invented content. No intro text. Bullets only.
                You MUST produce at least one bullet and NO MORE THAN 6 bullets total.
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
            [3-5 sentences of high-level professional prose for senior leadership. Focus on overall team progress and key accomplishments. Only mention a specific tower if it adds meaningful context — do not list or enumerate every tower. Do not mention risks, issues, or blockers. Do not name individuals.]
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

    private static string StripRiskLines(string text)
    {
        var lines = text.Split('\n');
        var kept = lines.Where(l =>
        {
            var lower = l.ToLowerInvariant();
            return !lower.Contains("risk") && !lower.Contains("blocker") && !lower.Contains("issue");
        });
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
