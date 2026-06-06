using System;
using System.IO;
using System.Text.Json;
using MailConvertPrivateUser.Services;
using Serilog;

namespace MailConvertPrivateUser;

public partial class App : System.Windows.Application
{
    protected override void OnStartup(System.Windows.StartupEventArgs e)
    {
        base.OnStartup(e);

        var logBaseDir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
        var eml2pstDir = Path.Combine(logBaseDir, "EML2PST");
        var pstDir = Path.Combine(logBaseDir, "PST");
        var ostDir = Path.Combine(logBaseDir, "OST");
        var imapDir = Path.Combine(logBaseDir, "IMAP");
        var o365Dir = Path.Combine(logBaseDir, "O365");

        Directory.CreateDirectory(eml2pstDir);
        Directory.CreateDirectory(pstDir);
        Directory.CreateDirectory(ostDir);
        Directory.CreateDirectory(imapDir);
        Directory.CreateDirectory(o365Dir);

        Log.Logger = new LoggerConfiguration()
            .MinimumLevel.Information()
            .WriteTo.File(Path.Combine(logBaseDir, "app-.log"), rollingInterval: RollingInterval.Day)
            .WriteTo.File(Path.Combine(eml2pstDir, "eml2pst-.log"), rollingInterval: RollingInterval.Day)
            .WriteTo.File(Path.Combine(pstDir, "pst-.log"), rollingInterval: RollingInterval.Day)
            .WriteTo.File(Path.Combine(ostDir, "ost-.log"), rollingInterval: RollingInterval.Day)
            .WriteTo.File(Path.Combine(imapDir, "imap-.log"), rollingInterval: RollingInterval.Day)
            .WriteTo.File(Path.Combine(o365Dir, "o365-.log"), rollingInterval: RollingInterval.Day)
            .CreateLogger();

        Log.Information("Application starting...");

        // Initialize localization from user settings
        InitializeLocalization();

        this.DispatcherUnhandledException += (s, args) =>
        {
            Log.Error(args.Exception, "Unhandled exception");
            var msg = LocalizationManager.GetString("App_StartupError", args.Exception.Message);
            System.Windows.MessageBox.Show(msg, LocalizationManager.GetString("Common_Error"),
                System.Windows.MessageBoxButton.OK, System.Windows.MessageBoxImage.Error);
            args.Handled = true;
        };
    }

    private static void InitializeLocalization()
    {
        try
        {
            var settingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "usersettings.json");
            string culture = "zh-CN";
            if (File.Exists(settingsPath))
            {
                var json = File.ReadAllText(settingsPath);
                var dict = JsonSerializer.Deserialize<System.Collections.Generic.Dictionary<string, string>>(json);
                if (dict != null && dict.TryGetValue("Language", out var lang) && !string.IsNullOrWhiteSpace(lang))
                {
                    culture = lang;
                }
            }
            LocalizationManager.Initialize(culture);
        }
        catch (Exception ex)
        {
            Log.Warning(ex, "Failed to initialize localization; defaulting to zh-CN");
            LocalizationManager.Initialize("zh-CN");
        }
    }

    protected override void OnExit(System.Windows.ExitEventArgs e)
    {
        Log.Information("Application exiting...");
        Log.CloseAndFlush();
        base.OnExit(e);
    }
}
