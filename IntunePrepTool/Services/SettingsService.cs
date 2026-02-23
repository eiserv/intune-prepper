using IntunePrepTool.Models;

namespace IntunePrepTool.Services;

internal static class SettingsService
{
    public static AppSettings Load()
    {
        string configPath = Path.Combine(AppContext.BaseDirectory, "config.yaml");
        AppSettings defaults = new() { ConfigPath = configPath };

        if (!File.Exists(configPath))
        {
            return defaults with
            {
                ConfigWarning = $"config.yaml not found at '{configPath}'. Built-in defaults are used."
            };
        }

        try
        {
            Dictionary<string, string> values = ParseSimpleYaml(configPath);

            return new AppSettings
            {
                AdminRecipientEmail = GetString(values, "admin_recipient_email", defaults.AdminRecipientEmail),
                OutputFolderName = GetString(values, "output_folder_name", defaults.OutputFolderName),
                IntuneAutopilotImportUrl = GetString(values, "intune_autopilot_import_url", defaults.IntuneAutopilotImportUrl),
                AutoEnrollCommand = GetString(values, "auto_enroll_command", defaults.AutoEnrollCommand),
                DefaultRunAutoEnroll = GetBool(values, "default_run_auto_enroll", defaults.DefaultRunAutoEnroll),
                DefaultAutoEnrollOnly = GetBool(values, "default_auto_enroll_only", defaults.DefaultAutoEnrollOnly),
                ConfigPath = configPath,
                ConfigWarning = string.Empty
            };
        }
        catch (Exception exception)
        {
            return defaults with
            {
                ConfigWarning = $"Could not parse config.yaml ({exception.Message}). Built-in defaults are used."
            };
        }
    }

    private static Dictionary<string, string> ParseSimpleYaml(string path)
    {
        Dictionary<string, string> values = new(StringComparer.OrdinalIgnoreCase);
        string[] lines = File.ReadAllLines(path);
        foreach (string rawLine in lines)
        {
            string line = rawLine.Trim();
            if (line.Length == 0 || line.StartsWith('#'))
            {
                continue;
            }

            int colonIndex = line.IndexOf(':');
            if (colonIndex <= 0)
            {
                continue;
            }

            string key = line[..colonIndex].Trim();
            string value = line[(colonIndex + 1)..].Trim();

            if (value.StartsWith('"') && value.EndsWith('"') && value.Length >= 2)
            {
                value = value[1..^1];
            }
            else if (value.StartsWith('\'') && value.EndsWith('\'') && value.Length >= 2)
            {
                value = value[1..^1];
            }

            values[key] = value;
        }

        return values;
    }

    private static string GetString(Dictionary<string, string> values, string key, string fallback)
    {
        return values.TryGetValue(key, out string? value) && !string.IsNullOrWhiteSpace(value)
            ? value.Trim()
            : fallback;
    }

    private static bool GetBool(Dictionary<string, string> values, string key, bool fallback)
    {
        return values.TryGetValue(key, out string? value) && bool.TryParse(value, out bool parsed)
            ? parsed
            : fallback;
    }
}
