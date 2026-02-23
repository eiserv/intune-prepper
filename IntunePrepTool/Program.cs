namespace IntunePrepTool;

static class Program
{
    [STAThread]
    static void Main()
    {
        ApplicationConfiguration.Initialize();
        var settings = Services.SettingsService.Load();
        Strings.SetLanguage(settings.Language);
        Application.Run(new MainForm(settings));
    }
}
