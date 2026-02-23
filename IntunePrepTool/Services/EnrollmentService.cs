using System.Diagnostics;
using System.Text.RegularExpressions;
using IntunePrepTool.Models;

namespace IntunePrepTool.Services;

internal sealed class EnrollmentService
{
    /// <summary>HRESULT 0x8018000A – device is already MDM-enrolled.</summary>
    private const int MENROLL_E_DEVICE_ALREADY_ENROLLED = unchecked((int)0x8018000A);

    private readonly PowerShellRunner _powerShellRunner = new();

    public async Task<EnrollmentAttemptResult> TryAutoEnrollAsync(
        DeviceCollectionResult deviceResult,
        string configuredCommand,
        CancellationToken cancellationToken = default)
    {
        bool azureAdJoined = MatchesYes(deviceResult.DsregStatus, "AzureAdJoined");
        bool mdmUrlPresent = HasMdmUrl(deviceResult.DsregStatus);

        if (!azureAdJoined)
        {
            return new EnrollmentAttemptResult
            {
                Attempted = false,
                AzureAdJoined = false,
                MdmUrlPresent = mdmUrlPresent,
                Command = configuredCommand,
                StatusMessage = Strings.EnrollSkippedNotJoined
            };
        }

        string command = Environment.ExpandEnvironmentVariables(configuredCommand);
        ProcessStartInfo startInfo = new()
        {
            FileName = "cmd.exe",
            Arguments = $"/c {command}",
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true
        };

        using Process process = new() { StartInfo = startInfo };
        process.Start();

        Task<string> outputTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        Task<string> errorTask = process.StandardError.ReadToEndAsync(cancellationToken);

        // Await streams and process exit together to avoid potential deadlocks.
        await Task.WhenAll(outputTask, errorTask, process.WaitForExitAsync(cancellationToken));

        string events = await QueryRecentEnrollmentEventsAsync(cancellationToken);
        int exitCode = process.ExitCode;
        bool alreadyEnrolled = exitCode == MENROLL_E_DEVICE_ALREADY_ENROLLED;
        bool succeeded = exitCode == 0 || alreadyEnrolled;

        string statusMessage = alreadyEnrolled
            ? Strings.EnrollAlreadyEnrolled
            : succeeded
                ? Strings.EnrollSucceeded
                : Strings.EnrollFailed;

        return new EnrollmentAttemptResult
        {
            Attempted = true,
            Succeeded = succeeded,
            ExitCode = exitCode,
            StandardOutput = (await outputTask).Trim(),
            StandardError = (await errorTask).Trim(),
            AzureAdJoined = azureAdJoined,
            MdmUrlPresent = mdmUrlPresent,
            Command = configuredCommand,
            StatusMessage = statusMessage,
            EnrollmentEvents = events
        };
    }

    private async Task<string> QueryRecentEnrollmentEventsAsync(CancellationToken cancellationToken)
    {
        string script = """
            $start = (Get-Date).AddMinutes(-15)
            $events = Get-WinEvent -FilterHashtable @{
                LogName = 'Microsoft-Windows-DeviceManagement-Enterprise-Diagnostics-Provider/Admin'
                StartTime = $start
            } -ErrorAction SilentlyContinue | Select-Object -First 10 TimeCreated, Id, LevelDisplayName, Message
            if ($null -eq $events) {
                'No recent MDM admin events found in the last 15 minutes.'
            }
            else {
                $events | Format-Table -Wrap | Out-String
            }
            """;

        PowerShellExecutionResult result = await _powerShellRunner.RunScriptAsync(script, cancellationToken);
        if (result.ExitCode != 0 && string.IsNullOrWhiteSpace(result.StandardOutput))
        {
            return string.IsNullOrWhiteSpace(result.StandardError)
                ? "Could not read recent MDM events."
                : $"Could not read recent MDM events: {result.StandardError}";
        }

        return string.IsNullOrWhiteSpace(result.StandardOutput)
            ? "No recent MDM admin events found in the last 15 minutes."
            : result.StandardOutput;
    }

    private static bool MatchesYes(string text, string key)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        Regex regex = new($@"^\s*{Regex.Escape(key)}\s*:\s*YES\s*$", RegexOptions.Multiline | RegexOptions.IgnoreCase);
        return regex.IsMatch(text);
    }

    private static bool HasMdmUrl(string text)
    {
        if (string.IsNullOrWhiteSpace(text))
        {
            return false;
        }

        Regex regex = new(@"^\s*MdmUrl\s*:\s*(.+?)\s*$", RegexOptions.Multiline | RegexOptions.IgnoreCase);
        Match match = regex.Match(text);
        if (!match.Success)
        {
            return false;
        }

        string value = match.Groups[1].Value.Trim();
        return !string.IsNullOrWhiteSpace(value) && !value.Equals("N/A", StringComparison.OrdinalIgnoreCase);
    }
}
