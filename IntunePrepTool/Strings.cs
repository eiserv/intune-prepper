namespace IntunePrepTool;

internal static class Strings
{
    private static string _language = "en";

    public static string Language => _language;

    public static void SetLanguage(string language)
    {
        _language = (language ?? "en").Trim().ToLowerInvariant() switch
        {
            "de" or "deutsch" or "german" => "de",
            _ => "en"
        };
    }

    // ── Window / general ────────────────────────────────────────────────
    public static string WindowTitle => L("Intune Prep Tool", "Intune Vorbereitungstool");

    public static string IntroText => L(
        "Use this tool to export Autopilot hash and optionally trigger MDM auto-enrollment.\n" +
        "Run as local administrator for reliable results.",
        "Dieses Tool exportiert den Autopilot-Hash und kann optional die MDM-Registrierung auslösen.\n" +
        "Für zuverlässige Ergebnisse als lokaler Administrator ausführen.");

    // ── Form labels ─────────────────────────────────────────────────────
    public static string GroupTag => L("Group Tag", "Gruppenkennzeichnung");
    public static string AssignedUserUpn => L("Assigned User (UPN)", "Zugewiesener Benutzer (UPN)");
    public static string AdminRecipient => L("Admin Recipient", "Admin-Empfänger");
    public static string WorkflowMode => L("Workflow Mode", "Arbeitsablauf");

    public static string AutoEnrollOnly => L(
        "Auto-enroll only (skip hash export)",
        "Nur automatisch registrieren (Hash-Export überspringen)");

    public static string AlsoRunAutoEnroll => L(
        "Also run auto-enrollment after export",
        "Auch automatische Registrierung nach Export ausführen");

    // ── Buttons ─────────────────────────────────────────────────────────
    public static string RunWorkflow => L("Run Workflow", "Workflow starten");
    public static string OpenOutputFolder => L("Open Output Folder", "Ausgabeordner öffnen");
    public static string OpenEmailDraft => L("Open Email Draft", "E-Mail-Entwurf öffnen");
    public static string OpenIntuneImportPage => L("Open Intune Import Page", "Intune-Importseite öffnen");
    public static string CopyMetadata => L("Copy Metadata", "Metadaten kopieren");

    // ── Status / result messages ────────────────────────────────────────
    public static string Ready => L("Ready.", "Bereit.");
    public static string NoRunYet => L("No run yet.", "Noch kein Durchlauf.");

    public static string ConfigLoaded(string path) =>
        L($"Loaded config from {path}", $"Konfiguration geladen aus {path}");

    public static string ConfigLoadedStatus(string path) =>
        L($"Config loaded from {path}", $"Konfiguration geladen aus {path}");

    public static string ValidationTitle => L("Validation", "Überprüfung");

    public static string ValidationGroupTagRequired => L(
        "Group Tag and Assigned User are required for hardware hash export mode.",
        "Gruppenkennzeichnung und zugewiesener Benutzer sind für den Hardware-Hash-Export erforderlich.");

    public static string WarningNotAdmin => L(
        "Warning: current user is not admin. Auto-enrollment command may fail.",
        "Warnung: Der aktuelle Benutzer ist kein Administrator. Der Registrierungsbefehl kann fehlschlagen.");

    public static string StartingWorkflow(bool autoEnrollOnly) =>
        L($"Starting workflow. Auto-enroll-only mode: {autoEnrollOnly}",
          $"Workflow wird gestartet. Nur-Registrierung-Modus: {autoEnrollOnly}");

    public static string HashExportSuccess(string csvPath) =>
        L($"Hardware hash export successful: {csvPath}",
          $"Hardware-Hash-Export erfolgreich: {csvPath}");

    public static string HashExportCompleted(string serial) =>
        L($"Hash export completed for serial '{serial}'.",
          $"Hash-Export für Seriennummer '{serial}' abgeschlossen.");

    public static string HashExportFailed(string reason) =>
        L($"Hardware hash export failed: {reason}",
          $"Hardware-Hash-Export fehlgeschlagen: {reason}");

    public static string HashExportFailedSummary => L(
        "Hardware hash export failed. Error report created.",
        "Hardware-Hash-Export fehlgeschlagen. Fehlerbericht erstellt.");

    public static string ErrorReportSaved(string path) =>
        L($"Error report saved to: {path}",
          $"Fehlerbericht gespeichert unter: {path}");

    public static string AutoEnrollOnlySummary => L(
        "Auto-enroll-only mode selected; hash export skipped.",
        "Nur-Registrierung-Modus gewählt; Hash-Export übersprungen.");

    public static string HashSkippedByMode => L(
        "Hardware hash step skipped due to selected workflow mode.",
        "Hardware-Hash-Schritt aufgrund des gewählten Arbeitsablaufs übersprungen.");

    public static string RunningAutoEnroll => L(
        "Running auto-enrollment command...",
        "Registrierungsbefehl wird ausgeführt…");

    public static string EnrollmentReportSaved(string path) =>
        L($"Enrollment report saved to: {path}",
          $"Registrierungsbericht gespeichert unter: {path}");

    public static string AutoEnrollStatus(string status) =>
        L($"Auto-enrollment status: {status}",
          $"Status der automatischen Registrierung: {status}");

    public static string WorkflowCompleteEnrolled => L(
        "Workflow completed. Auto-enrollment command executed.",
        "Workflow abgeschlossen. Registrierungsbefehl ausgeführt.");

    public static string WorkflowCompleteEnrollFailed => L(
        "Workflow completed, but auto-enrollment command reported failure.",
        "Workflow abgeschlossen, aber der Registrierungsbefehl hat einen Fehler gemeldet.");

    public static string WorkflowCompleteEnrollSkipped => L(
        "Workflow completed. Auto-enrollment was skipped based on join-state checks.",
        "Workflow abgeschlossen. Die Registrierung wurde aufgrund des Beitrittsstatus übersprungen.");

    public static string WorkflowCompleteAlreadyEnrolled => L(
        "Workflow completed. Device was already enrolled in MDM.",
        "Workflow abgeschlossen. Das Gerät war bereits in MDM registriert.");

    public static string WorkflowFailed => L(
        "Workflow failed unexpectedly.",
        "Workflow ist unerwartet fehlgeschlagen.");

    public static string UnexpectedError(string message) =>
        L($"Unexpected error: {message}",
          $"Unerwarteter Fehler: {message}");

    public static string ErrorTitle => L("Error", "Fehler");

    // ── Email ───────────────────────────────────────────────────────────
    public static string NoFilesForEmail => L(
        "No generated files found. Run workflow first.",
        "Keine erzeugten Dateien gefunden. Workflow zuerst starten.");

    public static string EmailDraftOutlook => L(
        "Email draft opened in Outlook with attachments.",
        "E-Mail-Entwurf in Outlook mit Anhängen geöffnet.");

    public static string EmailDraftDefault => L(
        "Email draft opened in default mail app. Attach generated files manually if needed.",
        "E-Mail-Entwurf in der Standard-Mail-App geöffnet. Dateien bei Bedarf manuell anhängen.");

    public static string EmailSubject(string deviceName, string serial) =>
        $"Intune Intake - {deviceName} - {serial}";

    public static string EmailBody(string deviceName, string serial, string user,
        bool hashCollected, bool enrollAttempted, bool enrollSuccess) =>
        L(
            $"Hello,{Environment.NewLine}{Environment.NewLine}" +
            $"please find the generated Intune prep files from this device." +
            $"{Environment.NewLine}{Environment.NewLine}" +
            $"Device: {deviceName}{Environment.NewLine}" +
            $"Serial: {serial}{Environment.NewLine}" +
            $"Collected by: {user}{Environment.NewLine}" +
            $"HashCollected: {hashCollected}{Environment.NewLine}" +
            $"AutoEnrollAttempted: {enrollAttempted}{Environment.NewLine}" +
            $"AutoEnrollSuccess: {enrollSuccess}{Environment.NewLine}" +
            $"{Environment.NewLine}Thanks.",
            $"Hallo,{Environment.NewLine}{Environment.NewLine}" +
            $"anbei die erzeugten Intune-Vorbereitungsdateien dieses Geräts." +
            $"{Environment.NewLine}{Environment.NewLine}" +
            $"Gerät: {deviceName}{Environment.NewLine}" +
            $"Seriennummer: {serial}{Environment.NewLine}" +
            $"Erfasst von: {user}{Environment.NewLine}" +
            $"Hash erfasst: {hashCollected}{Environment.NewLine}" +
            $"Registrierung versucht: {enrollAttempted}{Environment.NewLine}" +
            $"Registrierung erfolgreich: {enrollSuccess}{Environment.NewLine}" +
            $"{Environment.NewLine}Vielen Dank.");

    // ── Output folder ───────────────────────────────────────────────────
    public static string OutputFolderNotAvailable => L(
        "Output folder is not available yet.",
        "Ausgabeordner ist noch nicht verfügbar.");

    // ── Clipboard ───────────────────────────────────────────────────────
    public static string NoMetadataAvailable => L(
        "No metadata available yet.",
        "Noch keine Metadaten verfügbar.");

    public static string MetadataCopied => L(
        "Metadata copied to clipboard.",
        "Metadaten in die Zwischenablage kopiert.");

    // ── Enrollment messages ─────────────────────────────────────────────
    public static string EnrollSkippedNotJoined => L(
        "Skipped: device is not Azure AD joined (dsregcmd reports AzureAdJoined : NO).",
        "Übersprungen: Gerät ist nicht Azure AD beigetreten (dsregcmd meldet AzureAdJoined : NO).");

    public static string EnrollSucceeded => L(
        "Auto-enrollment command executed. Check Intune enrollment state and event log details.",
        "Registrierungsbefehl ausgeführt. Intune-Registrierungsstatus und Ereignisprotokoll prüfen.");

    public static string EnrollFailed => L(
        "Auto-enrollment command failed. Review stderr and event log snippets.",
        "Registrierungsbefehl fehlgeschlagen. Stderr und Ereignisprotokoll prüfen.");

    public static string EnrollAlreadyEnrolled => L(
        "Device is already enrolled in MDM. No action needed.",
        "Gerät ist bereits in MDM registriert. Keine Aktion erforderlich.");

    // ── Helper ──────────────────────────────────────────────────────────
    private static string L(string en, string de) => _language == "de" ? de : en;
}
