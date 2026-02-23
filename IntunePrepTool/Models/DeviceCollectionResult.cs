namespace IntunePrepTool.Models;

internal sealed class DeviceCollectionResult
{
    public string DeviceName { get; init; } = string.Empty;
    public string CurrentUser { get; init; } = string.Empty;
    public bool IsCurrentUserAdmin { get; init; }
    public string SerialNumber { get; init; } = string.Empty;
    public string Manufacturer { get; init; } = string.Empty;
    public string Model { get; init; } = string.Empty;
    public string BiosVersion { get; init; } = string.Empty;
    public string OsVersion { get; init; } = string.Empty;
    public string WindowsBuild { get; init; } = string.Empty;
    public string WindowsProductId { get; init; } = string.Empty;
    public string HardwareHash { get; init; } = string.Empty;
    public string HashError { get; init; } = string.Empty;
    public bool HashCollectionSkipped { get; init; }
    public string DsregStatus { get; init; } = string.Empty;
    public string TpmStatus { get; init; } = string.Empty;
    public DateTime TimestampUtc { get; init; } = DateTime.UtcNow;

    public bool HashCollectionSucceeded => !string.IsNullOrWhiteSpace(HardwareHash);
    public string FailureReason =>
        HashCollectionSkipped
            ? "Hardware hash collection was skipped by user selection."
            : HashCollectionSucceeded
            ? string.Empty
            : string.IsNullOrWhiteSpace(HashError)
                ? "Hardware hash value was empty."
                : HashError;
}
