using System.Text;
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
        string basePath = string.IsNullOrWhiteSpace(desktop) || !Directory.Exists(desktop)
            ? Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            : desktop;

        if (string.IsNullOrWhiteSpace(basePath))
        {
            basePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
        }

        string outputFolder = Path.Combine(basePath, _settings.OutputFolderName);
        Directory.CreateDirectory(outputFolder);
        return outputFolder;
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

        File.WriteAllText(reportPath, reportBuilder.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
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

        File.WriteAllText(reportPath, builder.ToString(), new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        return reportPath;
    }

    private static string EscapeCsv(string? value)
    {
        string normalized = value ?? string.Empty;
        return $"\"{normalized.Replace("\"", "\"\"")}\"";
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
