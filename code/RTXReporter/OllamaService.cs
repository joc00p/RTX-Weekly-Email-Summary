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

                Extract ALL specific activities, tasks, blockers, and next steps this person mentioned.
                Output as concise bullet points starting with -. No invented content. No intro text. Bullets only.
                You MUST produce at least one bullet. If details are sparse, summarize what little was provided.
                """;

            var summary = await CallOllama(prompt, ct);
            var trimmed = summary.Trim();
            // Always include this person — fall back to a raw excerpt if Ollama produced nothing usable
            if (string.IsNullOrWhiteSpace(trimmed) || trimmed.StartsWith("(no updates)", StringComparison.OrdinalIgnoreCase))
            {
                var fallbackLine = group.First().Subject;
                trimmed = $"- {(string.IsNullOrWhiteSpace(fallbackLine) ? "Updates submitted" : fallbackLine)}";
            }
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

        // Step 3: Generate summary + executive summary from the assembled sections
        StatusUpdate?.Invoke("Writing executive summary...");
        var execPrompt = $"""
            You are writing the final sections of a team status report for the period: {weekLabel}

            Individual team member updates:
            {teamSection}

            Write ONLY these two sections exactly as formatted:

            ### Summary
            - [overall team progress]
            - [key accomplishments]
            - [active risks or issues]
            - [items pending]

            ### Executive Summary
            [3-5 sentences of professional prose for senior leadership covering accomplishments, risks, and outlook]
            """;

        var execSummary = await CallOllama(execPrompt, ct);

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

        return CleanBulletSpacing(report.ToString());
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
