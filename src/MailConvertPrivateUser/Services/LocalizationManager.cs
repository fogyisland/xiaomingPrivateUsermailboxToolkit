using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;
using Serilog;

namespace MailConvertPrivateUser.Services;

public record LanguageInfo(string Code, string DisplayName);

public static class LocalizationManager
{
    private const string FallbackCulture = "zh-CN";
    private const string LogContext = "Localization";

    private static readonly Dictionary<string, string> _current = new();
    private static readonly Dictionary<string, string> _fallback = new();
    private static readonly HashSet<string> _warnedMissingKeys = new();
    private static readonly HashSet<string> _warnedFormatErrors = new();
    private static string _currentLanguage = FallbackCulture;
    private static string _languageDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Language");

    public static IReadOnlyList<LanguageInfo> AvailableLanguages { get; } = new List<LanguageInfo>
    {
        new("zh-CN", "中文（简体）"),
        new("en-US", "English")
    };

    public static string CurrentLanguage => _currentLanguage;

    // Test hooks
    public static void SetLanguageDirectoryForTesting(string dir) => _languageDirectory = dir;
    public static void ResetForTesting()
    {
        _current.Clear();
        _fallback.Clear();
        _warnedMissingKeys.Clear();
        _warnedFormatErrors.Clear();
        _currentLanguage = FallbackCulture;
        _languageDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Language");
    }

    /// <summary>
    /// Loads localized strings for the given culture. Must be called once on the UI thread
    /// during app startup, BEFORE any <see cref="GetString(string)"/> call. Not thread-safe;
    /// not designed for runtime re-init — language changes require app restart.
    /// </summary>
    /// <param name="culture">"zh-CN" or "en-US". Other values silently fall back to "zh-CN".</param>
    public static void Initialize(string culture)
    {
        _current.Clear();
        _fallback.Clear();

        if (string.IsNullOrWhiteSpace(culture) || !AvailableLanguages.Any(l => l.Code == culture))
        {
            Log.Warning("{Context}: unsupported culture '{Culture}', falling back to {Fallback}", LogContext, culture, FallbackCulture);
            culture = FallbackCulture;
        }

        _currentLanguage = culture;
        LoadInto(culture, _current);

        if (culture != FallbackCulture)
        {
            LoadInto(FallbackCulture, _fallback);
        }

        Log.Information("{Context}: initialized for {Culture} ({Count} keys loaded)", LogContext, culture, _current.Count);
    }

    public static string GetString(string key)
    {
        if (string.IsNullOrEmpty(key)) return string.Empty;
        if (_current.TryGetValue(key, out var v)) return v;
        if (_fallback.TryGetValue(key, out v)) return v;
        if (_warnedMissingKeys.Add(key))
        {
            Log.Warning("{Context}: missing key {Key} (culture: {Culture})", LogContext, key, _currentLanguage);
        }
        return $"[{key}]";
    }

    public static string GetString(string key, params object[] args)
    {
        var template = GetString(key);
        if (args == null || args.Length == 0) return template;
        try
        {
            return string.Format(template, args);
        }
        catch (FormatException ex)
        {
            if (_warnedFormatErrors.Add(key))
            {
                Log.Error(ex, "{Context}: format mismatch for key {Key} template='{Template}'", LogContext, key, template);
            }
            return template;
        }
    }

    private static void LoadInto(string culture, Dictionary<string, string> target)
    {
        var path = Path.Combine(_languageDirectory, culture.ToLowerInvariant() + ".xml");
        if (!File.Exists(path))
        {
            Log.Warning("{Context}: file not found at {Path}", LogContext, path);
            return;
        }
        try
        {
            var doc = XDocument.Load(path);
            var root = doc.Root;
            if (root == null) return;
            foreach (var s in root.Elements("string"))
            {
                var name = s.Attribute("name")?.Value;
                var value = s.Value;
                if (!string.IsNullOrEmpty(name) && value != null)
                {
                    target[name] = value;
                }
            }
        }
        catch (Exception ex)
        {
            Log.Error(ex, "{Context}: failed to load {Path}", LogContext, path);
        }
    }
}
