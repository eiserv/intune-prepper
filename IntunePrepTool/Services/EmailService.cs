using System.Diagnostics;
using System.Runtime.InteropServices;

namespace IntunePrepTool.Services;

internal sealed class EmailService
{
    /// <summary>
    /// Opens an email draft.
    /// When <paramref name="preferOutlookCom"/> is true, Outlook COM automation is
    /// tried first so that attachments can be added programmatically.
    /// Otherwise (default) the OS-default mail handler is used via the mailto protocol.
    /// </summary>
    /// <returns>true when Outlook COM was used (attachments included).</returns>
    public bool OpenDraft(string recipient, string subject, string body,
        IReadOnlyList<string> attachmentPaths, bool preferOutlookCom = false)
    {
        if (preferOutlookCom && attachmentPaths.Count > 0 &&
            TryOpenOutlookDraftWithAttachment(recipient, subject, body, attachmentPaths))
        {
            return true;
        }

        // Use mailto: – this opens whatever the user has set as default mail app.
        string mailto = BuildMailtoUri(recipient, subject, body);
        Process.Start(new ProcessStartInfo
        {
            FileName = mailto,
            UseShellExecute = true
        });

        return false;
    }

    private static bool TryOpenOutlookDraftWithAttachment(string recipient, string subject, string body, IReadOnlyList<string> attachmentPaths)
    {
        object? outlookApp = null;
        object? mailItem = null;

        try
        {
            Type? outlookType = Type.GetTypeFromProgID("Outlook.Application");
            if (outlookType is null)
            {
                return false;
            }

            outlookApp = Activator.CreateInstance(outlookType);
            if (outlookApp is null)
            {
                return false;
            }

            dynamic app = outlookApp;
            mailItem = app.CreateItem(0);
            dynamic mail = mailItem;
            mail.To = recipient;
            mail.Subject = subject;
            mail.Body = body;

            foreach (string attachmentPath in attachmentPaths)
            {
                if (File.Exists(attachmentPath))
                {
                    mail.Attachments.Add(attachmentPath);
                }
            }

            mail.Display(false);
            return true;
        }
        catch
        {
            return false;
        }
        finally
        {
            if (mailItem is not null)
            {
                Marshal.ReleaseComObject(mailItem);
            }

            if (outlookApp is not null)
            {
                Marshal.ReleaseComObject(outlookApp);
            }
        }
    }

    private static string BuildMailtoUri(string recipient, string subject, string body)
    {
        string escapedSubject = Uri.EscapeDataString(subject);
        string escapedBody = Uri.EscapeDataString(body);
        return $"mailto:{recipient}?subject={escapedSubject}&body={escapedBody}";
    }
}
