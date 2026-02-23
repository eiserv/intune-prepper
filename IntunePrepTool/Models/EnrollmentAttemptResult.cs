namespace IntunePrepTool.Models;

internal sealed class EnrollmentAttemptResult
{
    public bool Attempted { get; init; }
    public bool Succeeded { get; init; }
    public int ExitCode { get; init; }
    public string StandardOutput { get; init; } = string.Empty;
    public string StandardError { get; init; } = string.Empty;
    public bool AzureAdJoined { get; init; }
    public bool MdmUrlPresent { get; init; }
    public string Command { get; init; } = string.Empty;
    public string StatusMessage { get; init; } = string.Empty;
    public string EnrollmentEvents { get; init; } = string.Empty;
    public DateTime TimestampUtc { get; init; } = DateTime.UtcNow;
}
