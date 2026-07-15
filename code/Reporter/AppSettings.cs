using System;
using System.IO;
using System.Text.Json;

namespace Reporter;

public class AppSettings
{
    private static readonly string SettingsPath =
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData),
                     "RTXReporter", "settings.json");

    public string TemplatePath { get; set; } =
        Path.Combine(AppContext.BaseDirectory, "RTXReport", "RTX TEMPLATE.pptx");

    public static AppSettings Load()
    {
        try
        {
            if (File.Exists(SettingsPath))
            {
                var json = File.ReadAllText(SettingsPath);
                return JsonSerializer.Deserialize<AppSettings>(json) ?? new AppSettings();
            }
        }
        catch { }
        return new AppSettings();
    }

    public void Save()
    {
        Directory.CreateDirectory(Path.GetDirectoryName(SettingsPath)!);
        File.WriteAllText(SettingsPath, JsonSerializer.Serialize(this, new JsonSerializerOptions { WriteIndented = true }));
    }
}
