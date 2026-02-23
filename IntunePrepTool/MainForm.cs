using System.Diagnostics;
using IntunePrepTool.Models;
using IntunePrepTool.Services;

namespace IntunePrepTool;

internal sealed class MainForm : Form
{
    private readonly AppSettings _settings;
    private readonly DeviceCollector _collector = new();
    private readonly EnrollmentService _enrollmentService = new();
    private readonly ExportService _exportService;
    private readonly EmailService _emailService = new();

    private readonly TextBox _groupTagInput = new();
    private readonly TextBox _assignedUserInput = new();
    private readonly TextBox _statusOutput = new();
    private readonly Label _resultSummary = new();
    private readonly Label _configStatus = new();

    private readonly CheckBox _runAutoEnrollCheckbox = new();
    private readonly CheckBox _autoEnrollOnlyCheckbox = new();

    private readonly Button _runWorkflowButton = new();
    private readonly Button _openFolderButton = new();
    private readonly Button _mailButton = new();
    private readonly Button _openIntuneButton = new();
    private readonly Button _copyButton = new();

    private DeviceCollectionResult? _latestResult;
    private EnrollmentAttemptResult? _latestEnrollmentResult;
    private string? _latestOutputFolder;
    private readonly List<string> _latestAttachments = new();

    public MainForm(AppSettings settings)
    {
        _settings = settings;
        _exportService = new ExportService(_settings);
        InitializeUi();

        _runAutoEnrollCheckbox.Checked = _settings.DefaultRunAutoEnroll;
        _autoEnrollOnlyCheckbox.Checked = _settings.DefaultAutoEnrollOnly;
        ApplyModeToggle();

        if (!string.IsNullOrWhiteSpace(_settings.ConfigWarning))
        {
            _configStatus.Text = _settings.ConfigWarning;
            AppendStatus(_settings.ConfigWarning);
        }
        else
        {
            _configStatus.Text = $"Loaded config from {_settings.ConfigPath}";
            AppendStatus($"Config loaded from {_settings.ConfigPath}");
        }

        AppendStatus("Ready.");
    }

    private void InitializeUi()
    {
        Text = "Intune Prep Tool";
        StartPosition = FormStartPosition.CenterScreen;
        MinimumSize = new Size(980, 760);

        TableLayoutPanel root = new()
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            RowCount = 5,
            Padding = new Padding(16)
        };

        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.RowStyles.Add(new RowStyle(SizeType.Percent, 100f));
        root.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        Controls.Add(root);

        Label intro = new()
        {
            AutoSize = true,
            Text =
                "Use this tool to export Autopilot hash and optionally trigger MDM auto-enrollment.\n" +
                "Run as local administrator for reliable results.",
            MaximumSize = new Size(920, 0)
        };
        root.Controls.Add(intro, 0, 0);

        _configStatus.AutoSize = true;
        _configStatus.Margin = new Padding(0, 8, 0, 8);
        root.Controls.Add(_configStatus, 0, 1);

        TableLayoutPanel formPanel = new()
        {
            Dock = DockStyle.Top,
            ColumnCount = 2,
            AutoSize = true,
            Margin = new Padding(0, 8, 0, 12)
        };
        formPanel.ColumnStyles.Add(new ColumnStyle(SizeType.AutoSize));
        formPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100f));
        root.Controls.Add(formPanel, 0, 2);

        Label groupTagLabel = new() { Text = "Group Tag", AutoSize = true, Anchor = AnchorStyles.Left };
        _groupTagInput.Dock = DockStyle.Fill;
        _groupTagInput.Margin = new Padding(12, 4, 0, 8);

        Label assignedUserLabel = new() { Text = "Assigned User (UPN)", AutoSize = true, Anchor = AnchorStyles.Left };
        _assignedUserInput.Dock = DockStyle.Fill;
        _assignedUserInput.Margin = new Padding(12, 4, 0, 8);

        Label recipientLabel = new() { Text = "Admin Recipient", AutoSize = true, Anchor = AnchorStyles.Left };
        Label recipientValue = new()
        {
            Text = _settings.AdminRecipientEmail,
            AutoSize = true,
            Margin = new Padding(12, 8, 0, 8)
        };

        Label modeLabel = new() { Text = "Workflow Mode", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
        FlowLayoutPanel modePanel = new() { AutoSize = true, Margin = new Padding(8, 4, 0, 8) };

        _autoEnrollOnlyCheckbox.Text = "Auto-enroll only (skip hash export)";
        _autoEnrollOnlyCheckbox.AutoSize = true;
        _autoEnrollOnlyCheckbox.CheckedChanged += (_, _) => ApplyModeToggle();

        _runAutoEnrollCheckbox.Text = "Also run auto-enrollment after export";
        _runAutoEnrollCheckbox.AutoSize = true;

        modePanel.Controls.Add(_autoEnrollOnlyCheckbox);
        modePanel.Controls.Add(_runAutoEnrollCheckbox);

        formPanel.Controls.Add(groupTagLabel, 0, 0);
        formPanel.Controls.Add(_groupTagInput, 1, 0);
        formPanel.Controls.Add(assignedUserLabel, 0, 1);
        formPanel.Controls.Add(_assignedUserInput, 1, 1);
        formPanel.Controls.Add(recipientLabel, 0, 2);
        formPanel.Controls.Add(recipientValue, 1, 2);
        formPanel.Controls.Add(modeLabel, 0, 3);
        formPanel.Controls.Add(modePanel, 1, 3);

        _statusOutput.Dock = DockStyle.Fill;
        _statusOutput.Multiline = true;
        _statusOutput.ScrollBars = ScrollBars.Vertical;
        _statusOutput.ReadOnly = true;
        _statusOutput.BackColor = Color.White;
        _statusOutput.Font = new Font("Consolas", 10f, FontStyle.Regular);
        root.Controls.Add(_statusOutput, 0, 3);

        TableLayoutPanel bottom = new()
        {
            Dock = DockStyle.Fill,
            ColumnCount = 1,
            AutoSize = true
        };
        bottom.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        bottom.RowStyles.Add(new RowStyle(SizeType.AutoSize));
        root.Controls.Add(bottom, 0, 4);

        _resultSummary.AutoSize = true;
        _resultSummary.Text = "No run yet.";
        _resultSummary.Margin = new Padding(0, 0, 0, 8);
        bottom.Controls.Add(_resultSummary, 0, 0);

        FlowLayoutPanel buttonRow = new()
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = true
        };
        bottom.Controls.Add(buttonRow, 0, 1);

        _runWorkflowButton.Text = "Run Workflow";
        _runWorkflowButton.AutoSize = true;
        _runWorkflowButton.Click += async (_, _) => await RunWorkflowAsync();

        _openFolderButton.Text = "Open Output Folder";
        _openFolderButton.AutoSize = true;
        _openFolderButton.Enabled = false;
        _openFolderButton.Click += (_, _) => OpenOutputFolder();

        _mailButton.Text = "Open Email Draft";
        _mailButton.AutoSize = true;
        _mailButton.Enabled = false;
        _mailButton.Click += (_, _) => OpenEmailDraft();

        _openIntuneButton.Text = "Open Intune Import Page";
        _openIntuneButton.AutoSize = true;
        _openIntuneButton.Click += (_, _) => OpenUrl(_settings.IntuneAutopilotImportUrl);

        _copyButton.Text = "Copy Metadata";
        _copyButton.AutoSize = true;
        _copyButton.Enabled = false;
        _copyButton.Click += (_, _) => CopyMetadataToClipboard();

        buttonRow.Controls.Add(_runWorkflowButton);
        buttonRow.Controls.Add(_openFolderButton);
        buttonRow.Controls.Add(_mailButton);
        buttonRow.Controls.Add(_openIntuneButton);
        buttonRow.Controls.Add(_copyButton);
    }

    private void ApplyModeToggle()
    {
        bool autoEnrollOnly = _autoEnrollOnlyCheckbox.Checked;
        _groupTagInput.Enabled = !autoEnrollOnly;
        _assignedUserInput.Enabled = !autoEnrollOnly;

        if (autoEnrollOnly)
        {
            _runAutoEnrollCheckbox.Checked = true;
            _runAutoEnrollCheckbox.Enabled = false;
        }
        else
        {
            _runAutoEnrollCheckbox.Enabled = true;
        }
    }

    private async Task RunWorkflowAsync()
    {
        bool autoEnrollOnly = _autoEnrollOnlyCheckbox.Checked;
        string groupTag = _groupTagInput.Text.Trim();
        string assignedUser = _assignedUserInput.Text.Trim();
        bool runAutoEnroll = _runAutoEnrollCheckbox.Checked || autoEnrollOnly;

        if (!autoEnrollOnly && (string.IsNullOrWhiteSpace(groupTag) || string.IsNullOrWhiteSpace(assignedUser)))
        {
            MessageBox.Show(
                "Group Tag and Assigned User are required for hardware hash export mode.",
                "Validation",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        ToggleBusy(isBusy: true);
        _latestAttachments.Clear();
        _latestEnrollmentResult = null;
        AppendStatus($"Starting workflow. Auto-enroll-only mode: {autoEnrollOnly}");

        try
        {
            _latestResult = await _collector.CollectAsync(includeHardwareHash: !autoEnrollOnly);
            _latestOutputFolder = _exportService.EnsureOutputFolder();

            if (!_latestResult.IsCurrentUserAdmin)
            {
                AppendStatus("Warning: current user is not admin. Auto-enrollment command may fail.");
            }

            if (!autoEnrollOnly)
            {
                if (_latestResult.HashCollectionSucceeded)
                {
                    string csvPath = _exportService.WriteAutopilotCsv(_latestResult, groupTag, assignedUser, _latestOutputFolder);
                    _latestAttachments.Add(csvPath);
                    _resultSummary.Text = $"Hash export completed for serial '{Safe(_latestResult.SerialNumber)}'.";
                    AppendStatus($"Hardware hash export successful: {csvPath}");
                }
                else
                {
                    string errorReport = _exportService.WriteHashCollectionErrorReport(_latestResult, _latestOutputFolder);
                    _latestAttachments.Add(errorReport);
                    _resultSummary.Text = "Hardware hash export failed. Error report created.";
                    AppendStatus($"Hardware hash export failed: {Safe(_latestResult.FailureReason)}");
                    AppendStatus($"Error report saved to: {errorReport}");
                }
            }
            else
            {
                _resultSummary.Text = "Auto-enroll-only mode selected; hash export skipped.";
                AppendStatus("Hardware hash step skipped due to selected workflow mode.");
            }

            if (runAutoEnroll)
            {
                AppendStatus("Running auto-enrollment command...");
                _latestEnrollmentResult = await _enrollmentService.TryAutoEnrollAsync(_latestResult, _settings.AutoEnrollCommand);

                string enrollmentReport = _exportService.WriteEnrollmentReport(
                    _latestResult,
                    _latestEnrollmentResult,
                    _latestOutputFolder,
                    groupTag,
                    assignedUser);

                _latestAttachments.Add(enrollmentReport);
                AppendStatus($"Enrollment report saved to: {enrollmentReport}");
                AppendStatus($"Auto-enrollment status: {_latestEnrollmentResult.StatusMessage}");

                if (_latestEnrollmentResult.Attempted && _latestEnrollmentResult.Succeeded)
                {
                    _resultSummary.Text = "Workflow completed. Auto-enrollment command executed.";
                }
                else if (_latestEnrollmentResult.Attempted)
                {
                    _resultSummary.Text = "Workflow completed, but auto-enrollment command reported failure.";
                }
                else
                {
                    _resultSummary.Text = "Workflow completed. Auto-enrollment was skipped based on join-state checks.";
                }
            }

            EnableSecondaryActions();
            OpenOutputFolder();
            OpenEmailDraft();
        }
        catch (Exception exception)
        {
            _resultSummary.Text = "Workflow failed unexpectedly.";
            AppendStatus($"Unexpected error: {exception.Message}");
            MessageBox.Show(
                $"Unexpected error: {exception.Message}",
                "Error",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error);
        }
        finally
        {
            ToggleBusy(isBusy: false);
        }
    }

    private void OpenEmailDraft()
    {
        if (_latestResult is null || _latestAttachments.Count == 0)
        {
            AppendStatus("No generated files found. Run workflow first.");
            return;
        }

        string serial = Safe(_latestResult.SerialNumber);
        string subject = $"Intune Intake - {_latestResult.DeviceName} - {serial}";
        string body =
            $"Hello,{Environment.NewLine}{Environment.NewLine}" +
            "please find the generated Intune prep files from this device." +
            $"{Environment.NewLine}{Environment.NewLine}" +
            $"Device: {_latestResult.DeviceName}{Environment.NewLine}" +
            $"Serial: {serial}{Environment.NewLine}" +
            $"Collected by: {_latestResult.CurrentUser}{Environment.NewLine}" +
            $"HashCollected: {_latestResult.HashCollectionSucceeded}{Environment.NewLine}" +
            $"AutoEnrollAttempted: {_latestEnrollmentResult?.Attempted ?? false}{Environment.NewLine}" +
            $"AutoEnrollSuccess: {_latestEnrollmentResult?.Succeeded ?? false}{Environment.NewLine}" +
            $"{Environment.NewLine}Thanks.";

        bool usedOutlook = _emailService.OpenDraft(_settings.AdminRecipientEmail, subject, body, _latestAttachments);
        if (usedOutlook)
        {
            AppendStatus("Email draft opened in Outlook with attachments.");
        }
        else
        {
            AppendStatus("Email draft opened in default mail app. Attach generated files manually if needed.");
        }
    }

    private void OpenOutputFolder()
    {
        if (string.IsNullOrWhiteSpace(_latestOutputFolder) || !Directory.Exists(_latestOutputFolder))
        {
            AppendStatus("Output folder is not available yet.");
            return;
        }

        Process.Start(new ProcessStartInfo
        {
            FileName = _latestOutputFolder,
            UseShellExecute = true
        });
    }

    private void OpenUrl(string url)
    {
        Process.Start(new ProcessStartInfo
        {
            FileName = url,
            UseShellExecute = true
        });
    }

    private void CopyMetadataToClipboard()
    {
        if (_latestResult is null)
        {
            AppendStatus("No metadata available yet.");
            return;
        }

        string metadata =
            $"DeviceName: {_latestResult.DeviceName}{Environment.NewLine}" +
            $"SerialNumber: {Safe(_latestResult.SerialNumber)}{Environment.NewLine}" +
            $"GroupTag: {_groupTagInput.Text.Trim()}{Environment.NewLine}" +
            $"AssignedUser: {_assignedUserInput.Text.Trim()}{Environment.NewLine}" +
            $"HashCollected: {_latestResult.HashCollectionSucceeded}{Environment.NewLine}" +
            $"AutoEnrollAttempted: {_latestEnrollmentResult?.Attempted ?? false}{Environment.NewLine}" +
            $"AutoEnrollSucceeded: {_latestEnrollmentResult?.Succeeded ?? false}";

        Clipboard.SetText(metadata);
        AppendStatus("Metadata copied to clipboard.");
    }

    private void ToggleBusy(bool isBusy)
    {
        _runWorkflowButton.Enabled = !isBusy;
        Cursor = isBusy ? Cursors.WaitCursor : Cursors.Default;
    }

    private void EnableSecondaryActions()
    {
        _openFolderButton.Enabled = true;
        _mailButton.Enabled = true;
        _copyButton.Enabled = true;
    }

    private void AppendStatus(string message)
    {
        string line = $"[{DateTime.Now:HH:mm:ss}] {message}";
        if (_statusOutput.TextLength == 0)
        {
            _statusOutput.Text = line;
            return;
        }

        _statusOutput.AppendText(Environment.NewLine + line);
    }

    private static string Safe(string? value)
    {
        return string.IsNullOrWhiteSpace(value) ? "Unknown" : value;
    }
}
