using System.Text;
using System.Text.RegularExpressions;
using IntunePrepTool.Models;

namespace IntunePrepTool.Services;

internal sealed class ExportService
{
    private readonly AppSettings _settings;

    public ExportService(AppSettings settings)
    {
        _settings = settings;
    }

    public string EnsureOutputFolder()
    {
        string desktop = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
        if (string.IsNullOrWhiteSpace(desktop))
        {
            string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            desktop = string.IsNullOrWhiteSpace(userProfile)
                ? Path.Combine(Path.GetPathRoot(Environment.SystemDirectory) ?? "C:\\", "Users", "Default", "Desktop")
                : Path.Combine(userProfile, "Desktop");
        }

        try
        {
            Directory.CreateDirectory(desktop);
            string outputFolderOnDesktop = Path.Combine(desktop, _settings.OutputFolderName);
            Directory.CreateDirectory(outputFolderOnDesktop);
            return outputFolderOnDesktop;
        }
        catch
        {
            string fallbackBase = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (string.IsNullOrWhiteSpace(fallbackBase))
            {
                fallbackBase = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
            }

            if (string.IsNullOrWhiteSpace(fallbackBase))
            {
                fallbackBase = Path.GetTempPath();
            }

            string outputFolderFallback = Path.Combine(fallbackBase, _settings.OutputFolderName);
            Directory.CreateDirectory(outputFolderFallback);
            return outputFolderFallback;
        }
    }

    public string WriteAutopilotCsv(
        DeviceCollectionResult result,
        string groupTag,
        string assignedUser,
        string outputFolder)
    {
        string serial = SanitizeFilePart(result.SerialNumber, "UnknownSerial");
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmm");
        string filePath = Path.Combine(outputFolder, $"Autopilot_{serial}_{timestamp}.csv");

        StringBuilder csvBuilder = new();
        csvBuilder.AppendLine("Device Serial Number,Windows Product ID,Hardware Hash,Group Tag,Assigned User");
        csvBuilder.AppendLine(string.Join(",",
            EscapeCsv(result.SerialNumber),
            EscapeCsv(result.WindowsProductId),
            EscapeCsv(result.HardwareHash),
            EscapeCsv(groupTag),
            EscapeCsv(assignedUser)));

        File.WriteAllText(filePath, csvBuilder.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        return filePath;
    }

    public string WriteEnrollmentReport(
        DeviceCollectionResult result,
        EnrollmentAttemptResult enrollmentResult,
        string outputFolder,
        string groupTag,
        string assignedUser)
    {
        string serial = SanitizeFilePart(result.SerialNumber, "UnknownSerial");
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmm");
        string reportPath = Path.Combine(outputFolder, $"EnrollmentReport_{serial}_{timestamp}.txt");

        StringBuilder reportBuilder = new();
        reportBuilder.AppendLine("Intune Prep Tool - Auto-Enrollment Report");
        reportBuilder.AppendLine($"Generated (local): {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        reportBuilder.AppendLine($"Generated (UTC): {enrollmentResult.TimestampUtc:O}");
        reportBuilder.AppendLine();
        reportBuilder.AppendLine($"DeviceName: {result.DeviceName}");
        reportBuilder.AppendLine($"CurrentUser: {result.CurrentUser}");
        reportBuilder.AppendLine($"IsCurrentUserAdmin: {result.IsCurrentUserAdmin}");
        reportBuilder.AppendLine($"SerialNumber: {result.SerialNumber}");
        reportBuilder.AppendLine($"GroupTagInput: {groupTag}");
        reportBuilder.AppendLine($"AssignedUserInput: {assignedUser}");
        reportBuilder.AppendLine();
        reportBuilder.AppendLine("AutoEnrollment");
        reportBuilder.AppendLine($"Attempted: {enrollmentResult.Attempted}");
        reportBuilder.AppendLine($"Succeeded: {enrollmentResult.Succeeded}");
        reportBuilder.AppendLine($"ExitCode: {enrollmentResult.ExitCode}");
        reportBuilder.AppendLine($"AzureAdJoined: {enrollmentResult.AzureAdJoined}");
        reportBuilder.AppendLine($"MdmUrlPresent: {enrollmentResult.MdmUrlPresent}");
        reportBuilder.AppendLine($"StatusMessage: {enrollmentResult.StatusMessage}");
        reportBuilder.AppendLine($"Command: {enrollmentResult.Command}");
        reportBuilder.AppendLine();
        reportBuilder.AppendLine("Command StdOut");
        reportBuilder.AppendLine(string.IsNullOrWhiteSpace(enrollmentResult.StandardOutput)
            ? "(empty)"
            : enrollmentResult.StandardOutput);
        reportBuilder.AppendLine();
        reportBuilder.AppendLine("Command StdErr");
        reportBuilder.AppendLine(string.IsNullOrWhiteSpace(enrollmentResult.StandardError)
            ? "(empty)"
            : enrollmentResult.StandardError);
        reportBuilder.AppendLine();
        reportBuilder.AppendLine("Recent MDM Events");
        reportBuilder.AppendLine(string.IsNullOrWhiteSpace(enrollmentResult.EnrollmentEvents)
            ? "(none)"
            : enrollmentResult.EnrollmentEvents);

        File.WriteAllText(reportPath, reportBuilder.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        return reportPath;
    }

    public string WriteIntuneReadinessReport(
        DeviceCollectionResult result,
        EnrollmentAttemptResult? enrollmentResult,
        string outputFolder,
        string groupTag,
        string assignedUser,
        bool autoEnrollOnly)
    {
        string serial = SanitizeFilePart(result.SerialNumber, "UnknownSerial");
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmm");
        string reportPath = Path.Combine(outputFolder, $"IntuneReadiness_{serial}_{timestamp}.txt");

        bool azureAdJoined = enrollmentResult?.AzureAdJoined ?? MatchesYes(result.DsregStatus, "AzureAdJoined");
        bool mdmUrlPresent = enrollmentResult?.MdmUrlPresent ?? HasMdmUrl(result.DsregStatus);
        bool hashCollected = result.HashCollectionSucceeded;
        bool enrollmentAttempted = enrollmentResult?.Attempted ?? false;
        bool enrollmentSucceeded = enrollmentResult?.Succeeded ?? false;
        bool alreadyEnrolled = enrollmentResult?.ExitCode == unchecked((int)0x8018000A);

        List<string> openTasks = new();
        if (!result.IsCurrentUserAdmin)
        {
            openTasks.Add("Run the tool as local administrator for reliable enrollment execution.");
        }

        if (!azureAdJoined)
        {
            openTasks.Add("Join the device to Microsoft Entra ID (Azure AD join) before Intune enrollment.");
        }

        if (!mdmUrlPresent)
        {
            openTasks.Add("Verify automatic MDM enrollment scope in Entra ID and that the user has an Intune license.");
        }

        if (!autoEnrollOnly && !hashCollected)
        {
            openTasks.Add("Collect hardware hash successfully and re-export the Autopilot CSV.");
        }

        if (!autoEnrollOnly && hashCollected)
        {
            openTasks.Add("Import the generated Autopilot CSV into Intune and wait for processing.");
        }

        if (!alreadyEnrolled && enrollmentAttempted && !enrollmentSucceeded)
        {
            openTasks.Add("Review enrollment report/event logs and retry auto-enrollment command.");
        }

        if (!enrollmentAttempted)
        {
            openTasks.Add("Run or enable auto-enrollment to trigger MDM enrollment from the device.");
        }

        openTasks.Add("In Intune, verify device appears, apply enrollment profile, and assign compliance/configuration policies.");

        StringBuilder builder = new();
        builder.AppendLine("Intune Prep Tool - Intune Readiness Report");
        builder.AppendLine($"Generated (local): {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        builder.AppendLine($"Generated (UTC): {DateTime.UtcNow:O}");
        builder.AppendLine();
        builder.AppendLine("Device");
        builder.AppendLine($"DeviceName: {result.DeviceName}");
        builder.AppendLine($"CurrentUser: {result.CurrentUser}");
        builder.AppendLine($"IsCurrentUserAdmin: {result.IsCurrentUserAdmin}");
        builder.AppendLine($"SerialNumber: {result.SerialNumber}");
        builder.AppendLine($"GroupTagInput: {groupTag}");
        builder.AppendLine($"AssignedUserInput: {assignedUser}");
        builder.AppendLine();
        builder.AppendLine("Readiness Checks");
        builder.AppendLine($"AzureAdJoined: {azureAdJoined}");
        builder.AppendLine($"MdmUrlPresent: {mdmUrlPresent}");
        builder.AppendLine($"HardwareHashCollected: {hashCollected}");
        builder.AppendLine($"AutoEnrollmentAttempted: {enrollmentAttempted}");
        builder.AppendLine($"AutoEnrollmentSucceeded: {enrollmentSucceeded}");
        builder.AppendLine($"AlreadyEnrolledDetected: {alreadyEnrolled}");
        builder.AppendLine();
        builder.AppendLine("What Is Still Required");
        if (openTasks.Count == 0)
        {
            builder.AppendLine("- No open tasks detected. Verify in Intune portal that device state is fully compliant.");
        }
        else
        {
            foreach (string task in openTasks.Distinct(StringComparer.OrdinalIgnoreCase))
            {
                builder.AppendLine($"- {task}");
            }
        }

        builder.AppendLine();
        builder.AppendLine("DSREGCMD /STATUS (raw)");
        builder.AppendLine(string.IsNullOrWhiteSpace(result.DsregStatus) ? "(empty)" : result.DsregStatus);

        File.WriteAllText(reportPath, builder.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        return reportPath;
    }

    public string WriteHashCollectionErrorReport(DeviceCollectionResult result, string outputFolder)
    {
        string serial = SanitizeFilePart(result.SerialNumber, "UnknownSerial");
        string timestamp = DateTime.Now.ToString("yyyyMMdd_HHmm");
        string reportPath = Path.Combine(outputFolder, $"HashCollectionError_{serial}_{timestamp}.txt");

        StringBuilder builder = new();
        builder.AppendLine("Intune Prep Tool - Hash Collection Error");
        builder.AppendLine($"Generated (local): {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
        builder.AppendLine($"Generated (UTC): {result.TimestampUtc:O}");
        builder.AppendLine();
        builder.AppendLine($"DeviceName: {result.DeviceName}");
        builder.AppendLine($"CurrentUser: {result.CurrentUser}");
        builder.AppendLine($"IsCurrentUserAdmin: {result.IsCurrentUserAdmin}");
        builder.AppendLine($"SerialNumber: {result.SerialNumber}");
        builder.AppendLine($"FailureReason: {result.FailureReason}");
        builder.AppendLine();
        builder.AppendLine("DSREGCMD /STATUS");
        builder.AppendLine(result.DsregStatus);

        File.WriteAllText(reportPath, builder.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        return reportPath;
    }

    private static string EscapeCsv(string? value)
    {
        string normalized = value ?? string.Empty;
        return $"\"{normalized.Replace("\"", "\"\"")}\"";
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

    private static string SanitizeFilePart(string? value, string fallback)
    {
        string source = string.IsNullOrWhiteSpace(value) ? fallback : value.Trim();
        char[] invalid = Path.GetInvalidFileNameChars();
        StringBuilder builder = new(source.Length);
        foreach (char character in source)
        {
            builder.Append(invalid.Contains(character) ? '_' : character);
        }

        return string.IsNullOrWhiteSpace(builder.ToString()) ? fallback : builder.ToString();
    }
}
