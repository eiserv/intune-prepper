namespace IntunePrepTool.Models;

internal sealed record AppSettings
{
    public string AdminRecipientEmail { get; init; } = "it-admin@company.com";
    public string OutputFolderName { get; init; } = "IntunePrepOutput";
    public string IntuneAutopilotImportUrl { get; init; } =
        "https://intune.microsoft.com/#view/Microsoft_Intune_Enrollment/AutopilotDevices.ReactView";
    public string AutoEnrollCommand { get; init; } = @"%windir%\system32\deviceenroller.exe /c /AutoEnrollMDM";
    public bool DefaultRunAutoEnroll { get; init; }
    public bool DefaultAutoEnrollOnly { get; init; }
    public string ConfigPath { get; init; } = string.Empty;
    public string ConfigWarning { get; init; } = string.Empty;
}
