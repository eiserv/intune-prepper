# Intune Prep Tool

Windows GUI tool for coworker-friendly Intune onboarding prep.

## Features
- Exports Autopilot hardware hash CSV (with Group Tag + Assigned User).
- Optional auto-enrollment trigger:
  - `%windir%\system32\deviceenroller.exe /c /AutoEnrollMDM`
- Auto-enroll-only mode (skips hash export).
- Saves files to `Desktop\IntunePrepOutput` (or `Documents\IntunePrepOutput` if desktop path is unavailable).
- Opens an email draft to the admin with generated attachments.
- Writes enrollment report with command result and recent MDM event log snippets.

## Configuration
Edit `IntunePrepTool\config.yaml`:

```yaml
admin_recipient_email: "it-admin@company.com"
output_folder_name: "IntunePrepOutput"
intune_autopilot_import_url: "https://intune.microsoft.com/#view/Microsoft_Intune_Enrollment/AutopilotDevices.ReactView"
auto_enroll_command: "%windir%\\system32\\deviceenroller.exe /c /AutoEnrollMDM"
default_run_auto_enroll: false
default_auto_enroll_only: false
```

## Build and publish
Run from repo root:

```powershell
.\publish.ps1
```

If self-contained publish fails because of restricted internet access:

```powershell
.\publish.ps1 -FrameworkDependent
```

Published executable path:

`dist\IntunePrepTool.exe`
