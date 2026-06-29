using System.Collections.Generic;
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
    private static readonly HttpClient Http = new() { Timeout = System.TimeSpan.FromMinutes(10) };

    public async Task<string> SummarizeWeekAsync(string weekLabel, List<EmailItem> emails, CancellationToken ct)
    {
        var blocks = new StringBuilder();
        for (int i = 0; i < emails.Count; i++)
        {
            var e = emails[i];
            var excerpt = e.Body.Length > 3000 ? e.Body[..3000] : e.Body;
            blocks.AppendLine($"--- Email {i + 1} ---");
            blocks.AppendLine($"From: {e.From}");
            blocks.AppendLine($"Subject: {e.Subject}");
            blocks.AppendLine($"Received: {e.Received}");
            blocks.AppendLine($"Body:\n{excerpt}");
            blocks.AppendLine();
        }

        var prompt = $"""
            You are a project manager assistant creating a weekly team status report.

            Week: {weekLabel}

            Source data:

            {blocks}

            Generate a structured status report using ONLY the following format. Do not mention emails, senders, punch lists, or that this came from emails. Write as if it is an official team report. Exclude any activities or updates from Joel Coopersmith — do not include him in the report at all.

            CRITICAL RULES:
            - Only include people who have actual, specific activities mentioned in the source data above
            - Do not invent, infer, or create placeholder sections for anyone
            - Each real person appears exactly once — consolidate all their activities across all weeks into one section
            - If a name has no concrete activities, omit them entirely
            - Do not include Joel Coopersmith under any circumstances

            ---

            ## STATUS REPORT — {weekLabel}

            ### Team Updates

            **[Full Name] — [Role or Area if clearly stated]**
            - [specific activity or task from the source data]
            - [specific activity or task from the source data]
            - [blocker or risk if mentioned]
            - [next step if mentioned]

            (Repeat only for people with real data — no empty sections, no placeholders)

            ---

            ### Summary
            - [overall team progress]
            - [key accomplishments]
            - [active risks or issues]
            - [items pending]

            ---

            ### Executive Summary
            Write 3-5 sentences in professional prose summarizing the overall state of the team and project for this period. Highlight the most important accomplishments, any critical risks or blockers, and the outlook going forward. This should read as a standalone paragraph suitable for senior leadership.

            ---

            Team Updates and Summary use bullet points only. Executive Summary is prose only. No invented content. Only facts from the source data.
            """;

        var payload = JsonSerializer.Serialize(new { model = Model, prompt, stream = false });
        var content = new StringContent(payload, Encoding.UTF8, "application/json");
        var response = await Http.PostAsync(Url, content, ct);
        response.EnsureSuccessStatusCode();
        var json = await response.Content.ReadAsStringAsync(ct);
        using var doc = JsonDocument.Parse(json);
        return doc.RootElement.GetProperty("response").GetString() ?? "";
    }
}
