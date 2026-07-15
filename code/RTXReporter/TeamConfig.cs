using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;

namespace RTXReporter;

public class TeamConfig
{
    private static readonly string ConfigPath = Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
        "RTXReporter", "teams.json");

    public static readonly string[] TowerNames = ["SAP", "DBA SQL", "CLOUD", "ITIL Svc Mgmt"];

    public Dictionary<string, List<string>> Teams { get; private set; }

    public TeamConfig()
    {
        Teams = LoadOrDefault();
    }

    private static Dictionary<string, List<string>> LoadOrDefault()
    {
        if (File.Exists(ConfigPath))
        {
            try
            {
                var json = File.ReadAllText(ConfigPath);
                var loaded = JsonSerializer.Deserialize<Dictionary<string, List<string>>>(json);
                if (loaded != null) return loaded;
            }
            catch { }
        }
        return DefaultTeams();
    }

    public void Save()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(ConfigPath)!);
        File.WriteAllText(ConfigPath, JsonSerializer.Serialize(Teams,
            new JsonSerializerOptions { WriteIndented = true }));
    }

    public string GetTeam(string senderName)
    {
        var lower = senderName.ToLowerInvariant();
        foreach (var (team, members) in Teams)
            foreach (var m in members)
            {
                var mLower = m.ToLowerInvariant();

                // High-confidence: the full member name appears in the sender string
                if (lower.Contains(mLower))
                    return team;

                // Otherwise require at least two of the member's name parts to appear
                // (handles "Last, First" vs "First Last" and nicknames, while preventing a
                // shared first name alone from bucketing a different person into this tower).
                var tokens = mLower
                    .Split(new[] { ' ', ',', '.' }, StringSplitOptions.RemoveEmptyEntries)
                    .Where(w => w.Length >= 2)
                    .ToList();
                if (tokens.Count == 0) continue;
                int need = Math.Min(2, tokens.Count);
                int hits = tokens.Count(w => lower.Contains(w));
                if (hits >= need)
                    return team;
            }
        return "Other";
    }

    public void AddMember(string tower, string name)
    {
        // Remove from any other tower first
        foreach (var members in Teams.Values)
            members.RemoveAll(m => string.Equals(m, name, StringComparison.OrdinalIgnoreCase));

        if (!Teams.ContainsKey(tower))
            Teams[tower] = new List<string>();
        if (!Teams[tower].Any(m => string.Equals(m, name, StringComparison.OrdinalIgnoreCase)))
            Teams[tower].Add(name);
    }

    public void RemoveMember(string tower, string name)
    {
        if (Teams.TryGetValue(tower, out var members))
            members.RemoveAll(m => string.Equals(m, name, StringComparison.OrdinalIgnoreCase));
    }

    private static Dictionary<string, List<string>> DefaultTeams() => new()
    {
        ["SAP"]           = new List<string> { "Nicholas Cottone", "Celix Velaides", "Joshua Deheer", "Alex Aguilera" },
        ["DBA SQL"]       = new List<string> { "Victor Olufosoye", "Josue Guerrero", "Devery Page" },
        ["CLOUD"]         = new List<string> { "Carlos Villa", "Christopher Monday", "Shawn Adams", "Tyler Yosick", "Chris McMullan", "Mark Leonhartsberger", "William Bishop", "Sala El Faiz", "Andres Corredor", "Ryan Herbert" },
        ["ITIL Svc Mgmt"] = new List<string> { "Matt Hamby" },
    };
}
