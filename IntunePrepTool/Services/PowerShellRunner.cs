using System.Diagnostics;
using System.Text;
using IntunePrepTool.Models;

namespace IntunePrepTool.Services;

internal sealed class PowerShellRunner
{
    public async Task<PowerShellExecutionResult> RunScriptAsync(string script, CancellationToken cancellationToken = default)
    {
        string encodedCommand = Convert.ToBase64String(Encoding.Unicode.GetBytes(script));

        ProcessStartInfo startInfo = new()
        {
            FileName = "powershell.exe",
            Arguments = $"-NoProfile -NonInteractive -ExecutionPolicy Bypass -EncodedCommand {encodedCommand}",
            UseShellExecute = false,
            CreateNoWindow = true,
            RedirectStandardOutput = true,
            RedirectStandardError = true,
            StandardOutputEncoding = Encoding.UTF8,
            StandardErrorEncoding = Encoding.UTF8
        };

        using Process process = new() { StartInfo = startInfo };
        process.Start();

        Task<string> stdOutTask = process.StandardOutput.ReadToEndAsync(cancellationToken);
        Task<string> stdErrTask = process.StandardError.ReadToEndAsync(cancellationToken);
        await process.WaitForExitAsync(cancellationToken);

        return new PowerShellExecutionResult
        {
            ExitCode = process.ExitCode,
            StandardOutput = (await stdOutTask).Trim(),
            StandardError = (await stdErrTask).Trim()
        };
    }
}
