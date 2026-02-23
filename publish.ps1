param(
    [string]$ProjectPath = ".\IntunePrepTool\IntunePrepTool.csproj",
    [string]$OutputPath = ".\dist",
    [switch]$FrameworkDependent,
    [ValidateSet("en", "de")]
    [string]$Language
)

$ErrorActionPreference = "Stop"
$env:DOTNET_CLI_HOME = Join-Path $PSScriptRoot ".dotnet_cli_home"

if ($FrameworkDependent) {
    dotnet publish $ProjectPath `
        -c Release `
        -o $OutputPath
}
else {
    dotnet publish $ProjectPath `
        -c Release `
        -r win-x64 `
        --self-contained true `
        /p:PublishSingleFile=true `
        /p:IncludeNativeLibrariesForSelfExtract=true `
        /p:EnableCompressionInSingleFile=true `
        -o $OutputPath
}

if ($LASTEXITCODE -ne 0) {
    throw "Publish failed. For self-contained output, make sure internet access to nuget.org is available."
}

# Patch the published config.yaml with the chosen language if specified.
if ($Language) {
    $configFile = Join-Path $OutputPath "config.yaml"
    if (Test-Path $configFile) {
        $content = Get-Content $configFile -Raw
        $content = $content -replace '(?m)^language:\s*"[^"]*"', "language: `"$Language`""
        Set-Content $configFile -Value $content -NoNewline
        Write-Host "Language set to '$Language' in published config.yaml"
    }
}

Write-Host ""
Write-Host "Publish complete. EXE path:"
Write-Host (Join-Path $OutputPath "IntunePrepTool.exe")
