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
            _configStatus.Text = Strings.ConfigLoaded(_settings.ConfigPath);
            AppendStatus(Strings.ConfigLoadedStatus(_settings.ConfigPath));
        }

        AppendStatus(Strings.Ready);
    }

    private void InitializeUi()
    {
        Text = Strings.WindowTitle;
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
            Text = Strings.IntroText,
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

        Label groupTagLabel = new() { Text = Strings.GroupTag, AutoSize = true, Anchor = AnchorStyles.Left };
        _groupTagInput.Dock = DockStyle.Fill;
        _groupTagInput.Margin = new Padding(12, 4, 0, 8);

        Label assignedUserLabel = new() { Text = Strings.AssignedUserUpn, AutoSize = true, Anchor = AnchorStyles.Left };
        _assignedUserInput.Dock = DockStyle.Fill;
        _assignedUserInput.Margin = new Padding(12, 4, 0, 8);

        Label recipientLabel = new() { Text = Strings.AdminRecipient, AutoSize = true, Anchor = AnchorStyles.Left };
        Label recipientValue = new()
        {
            Text = _settings.AdminRecipientEmail,
            AutoSize = true,
            Margin = new Padding(12, 8, 0, 8)
        };

        Label modeLabel = new() { Text = Strings.WorkflowMode, Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
        FlowLayoutPanel modePanel = new() { AutoSize = true, Margin = new Padding(8, 4, 0, 8) };

        _autoEnrollOnlyCheckbox.Text = Strings.AutoEnrollOnly;
        _autoEnrollOnlyCheckbox.AutoSize = true;
        _autoEnrollOnlyCheckbox.CheckedChanged += (_, _) => ApplyModeToggle();

        _runAutoEnrollCheckbox.Text = Strings.AlsoRunAutoEnroll;
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
        _resultSummary.Text = Strings.NoRunYet;
        _resultSummary.Margin = new Padding(0, 0, 0, 8);
        bottom.Controls.Add(_resultSummary, 0, 0);

        FlowLayoutPanel buttonRow = new()
        {
            Dock = DockStyle.Fill,
            AutoSize = true,
            WrapContents = true
        };
        bottom.Controls.Add(buttonRow, 0, 1);

        _runWorkflowButton.Text = Strings.RunWorkflow;
        _runWorkflowButton.AutoSize = true;
        _runWorkflowButton.Click += async (_, _) => await RunWorkflowAsync();

        _openFolderButton.Text = Strings.OpenOutputFolder;
        _openFolderButton.AutoSize = true;
        _openFolderButton.Enabled = false;
        _openFolderButton.Click += (_, _) => OpenOutputFolder();

        _mailButton.Text = Strings.OpenEmailDraft;
        _mailButton.AutoSize = true;
        _mailButton.Enabled = false;
        _mailButton.Click += (_, _) => OpenEmailDraft();

        _openIntuneButton.Text = Strings.OpenIntuneImportPage;
        _openIntuneButton.AutoSize = true;
        _openIntuneButton.Click += (_, _) => OpenUrl(_settings.IntuneAutopilotImportUrl);

        _copyButton.Text = Strings.CopyMetadata;
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
                Strings.ValidationGroupTagRequired,
                Strings.ValidationTitle,
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning);
            return;
        }

        ToggleBusy(isBusy: true);
        _latestAttachments.Clear();
        _latestEnrollmentResult = null;
        AppendStatus(Strings.StartingWorkflow(autoEnrollOnly));

        try
        {
            _latestResult = await _collector.CollectAsync(includeHardwareHash: !autoEnrollOnly);
            _latestOutputFolder = _exportService.EnsureOutputFolder();

            if (!_latestResult.IsCurrentUserAdmin)
            {
                AppendStatus(Strings.WarningNotAdmin);
            }

            if (!autoEnrollOnly)
            {
                if (_latestResult.HashCollectionSucceeded)
                {
                    string csvPath = _exportService.WriteAutopilotCsv(_latestResult, groupTag, assignedUser, _latestOutputFolder);
                    _latestAttachments.Add(csvPath);
                    _resultSummary.Text = Strings.HashExportCompleted(Safe(_latestResult.SerialNumber));
                    AppendStatus(Strings.HashExportSuccess(csvPath));
                }
                else
                {
                    string errorReport = _exportService.WriteHashCollectionErrorReport(_latestResult, _latestOutputFolder);
                    _latestAttachments.Add(errorReport);
                    _resultSummary.Text = Strings.HashExportFailedSummary;
                    AppendStatus(Strings.HashExportFailed(Safe(_latestResult.FailureReason)));
                    AppendStatus(Strings.ErrorReportSaved(errorReport));
                }
            }
            else
            {
                _resultSummary.Text = Strings.AutoEnrollOnlySummary;
                AppendStatus(Strings.HashSkippedByMode);
            }

            if (runAutoEnroll)
            {
                AppendStatus(Strings.RunningAutoEnroll);
                _latestEnrollmentResult = await _enrollmentService.TryAutoEnrollAsync(_latestResult, _settings.AutoEnrollCommand);

                string enrollmentReport = _exportService.WriteEnrollmentReport(
                    _latestResult,
                    _latestEnrollmentResult,
                    _latestOutputFolder,
                    groupTag,
                    assignedUser);

                _latestAttachments.Add(enrollmentReport);
                AppendStatus(Strings.EnrollmentReportSaved(enrollmentReport));
                AppendStatus(Strings.AutoEnrollStatus(_latestEnrollmentResult.StatusMessage));

                if (_latestEnrollmentResult.Attempted && _latestEnrollmentResult.Succeeded)
                {
                    _resultSummary.Text = _latestEnrollmentResult.ExitCode == unchecked((int)0x8018000A)
                        ? Strings.WorkflowCompleteAlreadyEnrolled
                        : Strings.WorkflowCompleteEnrolled;
                }
                else if (_latestEnrollmentResult.Attempted)
                {
                    _resultSummary.Text = Strings.WorkflowCompleteEnrollFailed;
                }
                else
                {
                    _resultSummary.Text = Strings.WorkflowCompleteEnrollSkipped;
                }
            }

            string readinessReport = _exportService.WriteIntuneReadinessReport(
                _latestResult,
                _latestEnrollmentResult,
                _latestOutputFolder,
                groupTag,
                assignedUser,
                autoEnrollOnly);
            _latestAttachments.Add(readinessReport);
            AppendStatus($"Intune readiness report saved to: {readinessReport}");

            EnableSecondaryActions();
            OpenOutputFolder();
            OpenEmailDraft();
        }
        catch (Exception exception)
        {
            _resultSummary.Text = Strings.WorkflowFailed;
            AppendStatus(Strings.UnexpectedError(exception.Message));
            MessageBox.Show(
                Strings.UnexpectedError(exception.Message),
                Strings.ErrorTitle,
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
            AppendStatus(Strings.NoFilesForEmail);
            return;
        }

        string serial = Safe(_latestResult.SerialNumber);
        string subject = Strings.EmailSubject(_latestResult.DeviceName, serial);
        string body = Strings.EmailBody(
            _latestResult.DeviceName,
            serial,
            _latestResult.CurrentUser,
            _latestResult.HardwareHash,
            _latestResult.HashCollectionSucceeded,
            _latestEnrollmentResult?.Attempted ?? false,
            _latestEnrollmentResult?.Succeeded ?? false);

        bool usedOutlook = _emailService.OpenDraft(
            _settings.AdminRecipientEmail, subject, body, _latestAttachments,
            _settings.PreferOutlookCom);
        if (usedOutlook)
        {
            AppendStatus(Strings.EmailDraftOutlook);
        }
        else
        {
            AppendStatus(Strings.EmailDraftDefault);
        }
    }

    private void OpenOutputFolder()
    {
        if (string.IsNullOrWhiteSpace(_latestOutputFolder) || !Directory.Exists(_latestOutputFolder))
        {
            AppendStatus(Strings.OutputFolderNotAvailable);
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
            AppendStatus(Strings.NoMetadataAvailable);
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
        AppendStatus(Strings.MetadataCopied);
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
