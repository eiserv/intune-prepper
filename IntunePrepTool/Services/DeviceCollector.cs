using System.Text.Json;
using IntunePrepTool.Models;

namespace IntunePrepTool.Services;

internal sealed class DeviceCollector
{
    private readonly PowerShellRunner _powerShellRunner = new();

    public async Task<DeviceCollectionResult> CollectAsync(
        bool includeHardwareHash = true,
        CancellationToken cancellationToken = default)
    {
        PowerShellExecutionResult executionResult =
            await _powerShellRunner.RunScriptAsync(GetCollectionScript(includeHardwareHash), cancellationToken);

        if (executionResult.ExitCode != 0)
        {
            return new DeviceCollectionResult
            {
                DeviceName = Environment.MachineName,
                CurrentUser = Environment.UserName,
                IsCurrentUserAdmin = false,
                HashError = string.IsNullOrWhiteSpace(executionResult.StandardError)
                    ? "PowerShell collection failed with unknown error."
                    : executionResult.StandardError
            };
        }

        if (string.IsNullOrWhiteSpace(executionResult.StandardOutput))
        {
            return new DeviceCollectionResult
            {
                DeviceName = Environment.MachineName,
                CurrentUser = Environment.UserName,
                IsCurrentUserAdmin = false,
                HashError = "PowerShell did not return any output."
            };
        }

        try
        {
            using JsonDocument jsonDocument = JsonDocument.Parse(executionResult.StandardOutput);
            JsonElement root = jsonDocument.RootElement;

            return new DeviceCollectionResult
            {
                DeviceName = GetString(root, "DeviceName"),
                CurrentUser = GetString(root, "CurrentUser"),
                IsCurrentUserAdmin = GetBoolean(root, "IsAdmin"),
                SerialNumber = GetString(root, "SerialNumber"),
                Manufacturer = GetString(root, "Manufacturer"),
                Model = GetString(root, "Model"),
                BiosVersion = GetString(root, "BiosVersion"),
                OsVersion = GetString(root, "OsVersion"),
                WindowsBuild = GetString(root, "WindowsBuild"),
                WindowsProductId = GetString(root, "WindowsProductId"),
                HardwareHash = GetString(root, "HardwareHash"),
                HashError = GetString(root, "HashError"),
                HashCollectionSkipped = GetBoolean(root, "HashCollectionSkipped"),
                DsregStatus = GetString(root, "DsregStatus"),
                TpmStatus = GetString(root, "TpmStatus"),
                TimestampUtc = ParseTimestamp(root)
            };
        }
        catch (Exception exception)
        {
            return new DeviceCollectionResult
            {
                DeviceName = Environment.MachineName,
                CurrentUser = Environment.UserName,
                IsCurrentUserAdmin = false,
                HashError = $"Unable to parse collection output: {exception.Message}"
            };
        }
    }

    private static string GetCollectionScript(bool includeHardwareHash)
    {
        string includeHashFlag = includeHardwareHash ? "$true" : "$false";

        string script = """
            $ErrorActionPreference = 'Stop'
            $includeHardwareHash = __INCLUDE_HASH__

            $isAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).
                IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)

            $bios = Get-CimInstance -ClassName Win32_BIOS
            $cs = Get-CimInstance -ClassName Win32_ComputerSystem
            $os = Get-CimInstance -ClassName Win32_OperatingSystem

            try {
                $product = Get-ItemProperty -Path 'HKLM:\SOFTWARE\Microsoft\Windows NT\CurrentVersion' -ErrorAction Stop
                $productId = [string]$product.ProductId
            }
            catch {
                $productId = ''
            }

            $hash = ''
            $hashError = ''
            $hashCollectionSkipped = $false
            if ($includeHardwareHash) {
                try {
                    $hashResult = Get-CimInstance `
                        -Namespace root/cimv2/mdm/dmmap `
                        -ClassName MDM_DevDetail_Ext01 `
                        -Filter "InstanceID='Ext' AND ParentID='./DevDetail'" `
                        -ErrorAction Stop

                    $hash = [string]$hashResult.DeviceHardwareData
                    if ([string]::IsNullOrWhiteSpace($hash)) {
                        $hashError = 'Hardware hash query returned an empty value.'
                    }
                }
                catch {
                    $hashError = $_.Exception.Message
                }
            }
            else {
                $hashCollectionSkipped = $true
            }

            try {
                $dsreg = (dsregcmd /status | Out-String).Trim()
            }
            catch {
                $dsreg = "Unavailable: $($_.Exception.Message)"
            }

            try {
                $tpm = Get-Tpm -ErrorAction Stop | ConvertTo-Json -Compress
            }
            catch {
                $tpm = "Unavailable: $($_.Exception.Message)"
            }

            [pscustomobject]@{
                DeviceName = $env:COMPUTERNAME
                CurrentUser = "$env:USERDOMAIN\$env:USERNAME"
                IsAdmin = [bool]$isAdmin
                SerialNumber = [string]$bios.SerialNumber
                Manufacturer = [string]$cs.Manufacturer
                Model = [string]$cs.Model
                BiosVersion = [string]$bios.SMBIOSBIOSVersion
                OsVersion = [string]$os.Version
                WindowsBuild = [string]$os.BuildNumber
                WindowsProductId = [string]$productId
                HardwareHash = [string]$hash
                HashError = [string]$hashError
                HashCollectionSkipped = [bool]$hashCollectionSkipped
                DsregStatus = [string]$dsreg
                TpmStatus = [string]$tpm
                TimestampUtc = (Get-Date).ToUniversalTime().ToString('o')
            } | ConvertTo-Json -Compress -Depth 5
            """;

        return script.Replace("__INCLUDE_HASH__", includeHashFlag, StringComparison.Ordinal);
    }

    private static string GetString(JsonElement root, string propertyName)
    {
        if (!root.TryGetProperty(propertyName, out JsonElement element))
        {
            return string.Empty;
        }

        return element.GetString() ?? string.Empty;
    }

    private static bool GetBoolean(JsonElement root, string propertyName)
    {
        if (!root.TryGetProperty(propertyName, out JsonElement element))
        {
            return false;
        }

        return element.ValueKind switch
        {
            JsonValueKind.True => true,
            JsonValueKind.False => false,
            JsonValueKind.String when bool.TryParse(element.GetString(), out bool value) => value,
            _ => false
        };
    }

    private static DateTime ParseTimestamp(JsonElement root)
    {
        string value = GetString(root, "TimestampUtc");
        return DateTime.TryParse(value, out DateTime timestamp) ? timestamp.ToUniversalTime() : DateTime.UtcNow;
    }
}
