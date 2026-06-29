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

            Generate a structured weekly status report using ONLY the following format. Do not mention emails, senders, punch lists, or that this came from emails. Write as if it is an official team report. Exclude any activities or updates from Joel Coopersmith — do not include him in the report at all.

            ---

            ## WEEKLY STATUS REPORT — {weekLabel}

            ### Team Updates
            For each team member's work, create a section:

            **[Full Name] — [Role or Area if inferable]**
            - [bullet: key activity or task completed]
            - [bullet: key activity or task completed]
            - [bullet: issues, blockers, or risks if any]
            - [bullet: pending actions or next steps if any]

            (Repeat for each person)

            ---

            ### Weekly Summary
            - [bullet: overall team progress this week]
            - [bullet: key accomplishments]
            - [bullet: active risks or issues]
            - [bullet: items pending or in progress]

            ---

            Use consistent formatting. Use bullet points only — no paragraphs. Be concise and factual.
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
