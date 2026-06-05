# i18n Language Switching Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add Chinese / English language switching to LiteEMLTOPST. Users pick a language in the existing PreferencesWindow; the entire UI is translated after restart. All UI text lives in external XML files under `Language/`.

**Architecture:** A custom `LocalizationManager` (loads `Language/{culture}.xml` into a `Dictionary<string,string>`) plus a WPF `MarkupExtension` (`LocExtension`) that lets XAML bind with `{loc:Loc Key}`. Lookup order on miss: current culture → zh-CN fallback → return `[key]`. Language choice is persisted to `usersettings.json["Language"]`; restart is required to apply (no hot-reload).

**Tech Stack:** WPF, .NET 8, xunit, Serilog, System.Text.Json, System.Xml.Linq

**Translation catalog:** `docs/superpowers/translations/keys.csv` is the single source of truth for all `key → zh-CN → en-US` mappings used in this plan. Tasks reference keys by name.

**Reference spec:** `docs/superpowers/specs/2026-06-05-i18n-language-switching-design.md`

---

## File Structure

**Create:**
- `Language/zh-cn.xml` — Chinese strings (default + fallback)
- `Language/en-us.xml` — English strings
- `src/MailConvertPrivateUser/Services/LocalizationManager.cs` — load + lookup
- `src/MailConvertPrivateUser/Markup/LocExtension.cs` — WPF markup extension
- `tests/LiteEMLTOPST.Tests/Services/LocalizationManagerTests.cs` — xunit tests

**Modify:**
- `src/MailConvertPrivateUser/MailConvertPrivateUser.csproj` — add `<None Include="Language\*.xml">` to copy XMLs to output
- `src/MailConvertPrivateUser/App.xaml.cs` — call `LocalizationManager.Initialize` on startup; translate hardcoded "错误" string
- `src/MailConvertPrivateUser/PreferencesWindow.xaml` — add `xmlns:loc`; add `Language` Tab; replace hardcoded text
- `src/MailConvertPrivateUser/PreferencesWindow.xaml.cs` — load/save `Language` field; restart prompt
- `src/MailConvertPrivateUser/ImapAccountEditWindow.xaml` — add `xmlns:loc`; replace text
- `src/MailConvertPrivateUser/O365AccountEditWindow.xaml` — add `xmlns:loc`; replace text
- `src/MailConvertPrivateUser/ContactMappingDialog.xaml` — add `xmlns:loc`; replace text
- `src/MailConvertPrivateUser/RegistrationWindow.xaml` — add `xmlns:loc`; replace text
- `src/MailConvertPrivateUser/RegistrationWindow.xaml.cs` — translate MessageBox strings
- `src/MailConvertPrivateUser/MainWindow.xaml` — add `xmlns:loc`; replace text (largest file, ~226 strings)
- `src/MailConvertPrivateUser/MainWindow.xaml.cs` — translate MessageBox strings

---

## Task 1: Create translation catalog

**Files:**
- Create: `Language/zh-cn.xml` (skeleton)
- Create: `Language/en-us.xml` (skeleton)

- [ ] **Step 1: Create `Language/` directory and the two XML files with skeleton structure**

`Language/zh-cn.xml`:
```xml
<?xml version="1.0" encoding="utf-8"?>
<resources>
  <string name="Common_OK">确定</string>
  <string name="Common_Cancel">取消</string>
  <!-- All keys from docs/superpowers/translations/keys.csv will be added in Task 3 -->
</resources>
```

`Language/en-us.xml`:
```xml
<?xml version="1.0" encoding="utf-8"?>
<resources>
  <string name="Common_OK">OK</string>
  <string name="Common_Cancel">Cancel</string>
  <!-- All keys from docs/superpowers/translations/keys.csv will be added in Task 3 -->
</resources>
```

- [ ] **Step 2: Commit**

```bash
cd D:/ToolDevelop/LiteEMLTOPST
git add Language/
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -m "feat(i18n): add Language directory with skeleton XML files"
```

---

## Task 2: Configure csproj to copy Language files to output

**Files:**
- Modify: `src/MailConvertPrivateUser/MailConvertPrivateUser.csproj:19-21`

- [ ] **Step 1: Add the Language XML files to the project's copy-to-output item group**

In `src/MailConvertPrivateUser/MailConvertPrivateUser.csproj`, find the existing `<ItemGroup>` that contains `<None Update="appsettings.json">`. After the `<Content Include="app.ico">` element, add a new line:

```xml
    <None Update="Language\*.xml">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </None>
```

The full ItemGroup should look like:
```xml
  <ItemGroup>
    <None Update="appsettings.json">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </None>
    <Content Include="app.ico">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </Content>
    <None Update="Language\*.xml">
      <CopyToOutputDirectory>PreserveNewest</CopyToOutputDirectory>
    </None>
  </ItemGroup>
```

- [ ] **Step 2: Verify build still works**

Run: `cd D:/ToolDevelop/LiteEMLTOPST/src/MailConvertPrivateUser && dotnet build 2>&1 | tail -20`
Expected: `Build succeeded.` (warnings OK, no errors)

- [ ] **Step 3: Commit**

```bash
cd D:/ToolDevelop/LiteEMLTOPST
git add src/MailConvertPrivateUser/MailConvertPrivateUser.csproj
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -m "build: copy Language xml files to output directory"
```

---

## Task 3: Implement LocalizationManager — write failing tests first (TDD)

**Files:**
- Create: `tests/LiteEMLTOPST.Tests/Services/LocalizationManagerTests.cs`

- [ ] **Step 1: Write the failing test file**

`tests/LiteEMLTOPST.Tests/Services/LocalizationManagerTests.cs`:
```csharp
using System;
using System.IO;
using Xunit;
using MailConvertPrivateUser.Services;

namespace LiteEMLTOPST.Tests.Services;

[Collection("LocalizationManagerTests")]
public class LocalizationManagerTests : IDisposable
{
    private readonly string _testLangDir;

    public LocalizationManagerTests()
    {
        _testLangDir = Path.Combine(Path.GetTempPath(), "LiteEMLTOPST_TestLang_" + Guid.NewGuid());
        Directory.CreateDirectory(_testLangDir);
        LocalizationManager.SetLanguageDirectoryForTesting(_testLangDir);
    }

    public void Dispose()
    {
        if (Directory.Exists(_testLangDir))
        {
            Directory.Delete(_testLangDir, recursive: true);
        }
        LocalizationManager.ResetForTesting();
    }

    private void WriteXml(string fileName, string content)
    {
        File.WriteAllText(Path.Combine(_testLangDir, fileName), content);
    }

    [Fact]
    public void GetString_ReturnsChinese_WhenInitializedWithZhCN()
    {
        WriteXml("zh-cn.xml", "<?xml version=\"1.0\"?><resources><string name=\"Greeting\">你好</string></resources>");
        WriteXml("en-us.xml", "<?xml version=\"1.0\"?><resources><string name=\"Greeting\">Hello</string></resources>");

        LocalizationManager.Initialize("zh-CN");

        Assert.Equal("你好", LocalizationManager.GetString("Greeting"));
    }

    [Fact]
    public void GetString_ReturnsEnglish_WhenInitializedWithEnUS()
    {
        WriteXml("zh-cn.xml", "<?xml version=\"1.0\"?><resources><string name=\"Greeting\">你好</string></resources>");
        WriteXml("en-us.xml", "<?xml version=\"1.0\"?><resources><string name=\"Greeting\">Hello</string></resources>");

        LocalizationManager.Initialize("en-US");

        Assert.Equal("Hello", LocalizationManager.GetString("Greeting"));
    }

    [Fact]
    public void GetString_FallsBackToChinese_WhenKeyMissingInCurrentCulture()
    {
        WriteXml("zh-cn.xml", "<?xml version=\"1.0\"?><resources><string name=\"OnlyInChinese\">只在中文</string></resources>");
        WriteXml("en-us.xml", "<?xml version=\"1.0\"?><resources><string name=\"Greeting\">Hello</string></resources>");

        LocalizationManager.Initialize("en-US");

        Assert.Equal("只在中文", LocalizationManager.GetString("OnlyInChinese"));
    }

    [Fact]
    public void GetString_ReturnsBracketedKey_WhenKeyMissingInBothCultures()
    {
        WriteXml("zh-cn.xml", "<?xml version=\"1.0\"?><resources><string name=\"Other\">其他</string></resources>");
        WriteXml("en-us.xml", "<?xml version=\"1.0\"?><resources><string name=\"Greeting\">Hello</string></resources>");

        LocalizationManager.Initialize("en-US");

        Assert.Equal("[MissingKey]", LocalizationManager.GetString("MissingKey"));
    }

    [Fact]
    public void GetString_SupportsFormatPlaceholders()
    {
        WriteXml("zh-cn.xml", "<?xml version=\"1.0\"?><resources><string name=\"Welcome\">欢迎 {0}</string></resources>");
        WriteXml("en-us.xml", "<?xml version=\"1.0\"?><resources><string name=\"Welcome\">Welcome {0}</string></resources>");

        LocalizationManager.Initialize("en-US");

        Assert.Equal("Welcome Alice", LocalizationManager.GetString("Welcome", "Alice"));
    }

    [Fact]
    public void Initialize_DefaultsToChinese_WhenCultureFileMissing()
    {
        WriteXml("zh-cn.xml", "<?xml version=\"1.0\"?><resources><string name=\"Greeting\">你好</string></resources>");
        // no en-us.xml

        LocalizationManager.Initialize("en-US");

        Assert.Equal("你好", LocalizationManager.GetString("Greeting"));
    }

    [Fact]
    public void Initialize_NeverThrows_WhenAllFilesMissing()
    {
        // directory is empty
        var ex = Record.Exception(() => LocalizationManager.Initialize("en-US"));
        Assert.Null(ex);
        Assert.Equal("[Anything]", LocalizationManager.GetString("Anything"));
    }

    [Fact]
    public void CurrentLanguage_ReflectsInitialization()
    {
        WriteXml("zh-cn.xml", "<?xml version=\"1.0\"?><resources></resources>");
        WriteXml("en-us.xml", "<?xml version=\"1.0\"?><resources></resources>");

        LocalizationManager.Initialize("en-US");

        Assert.Equal("en-US", LocalizationManager.CurrentLanguage);
    }

    [Fact]
    public void AvailableLanguages_ContainsBothSupportedCultures()
    {
        var langs = LocalizationManager.AvailableLanguages;
        Assert.Contains(langs, l => l.Code == "zh-CN");
        Assert.Contains(langs, l => l.Code == "en-US");
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `cd D:/ToolDevelop/LiteEMLTOPST && dotnet test tests/LiteEMLTOPST.Tests/LiteEMLTOPST.Tests.csproj 2>&1 | tail -15`
Expected: `Build FAILED` with errors like `The type or namespace name 'LocalizationManager' could not be found` (we haven't created the class yet).

---

## Task 4: Implement LocalizationManager to make tests pass

**Files:**
- Create: `src/MailConvertPrivateUser/Services/LocalizationManager.cs`

- [ ] **Step 1: Create the LocalizationManager class**

`src/MailConvertPrivateUser/Services/LocalizationManager.cs`:
```csharp
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
        _currentLanguage = FallbackCulture;
        _languageDirectory = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Language");
    }

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
        Log.Warning("{Context}: missing key {Key} (culture: {Culture})", LogContext, key, _currentLanguage);
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
        catch (FormatException)
        {
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
```

- [ ] **Step 2: Run tests to verify they pass**

Run: `cd D:/ToolDevelop/LiteEMLTOPST && dotnet test tests/LiteEMLTOPST.Tests/LiteEMLTOPST.Tests.csproj 2>&1 | tail -20`
Expected: `Passed! - Failed: 0, Passed: 9`

- [ ] **Step 3: Commit**

```bash
cd D:/ToolDevelop/LiteEMLTOPST
git add src/MailConvertPrivateUser/Services/LocalizationManager.cs tests/LiteEMLTOPST.Tests/Services/LocalizationManagerTests.cs
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -m "feat(i18n): add LocalizationManager with xunit tests"
```

---

## Task 5: Create LocExtension

**Files:**
- Create: `src/MailDevelop\LiteEMLTOPST\src\MailConvertPrivateUser\Markup\LocExtension.cs`

- [ ] **Step 1: Create the markup extension**

`src/MailConvertPrivateUser/Markup/LocExtension.cs`:
```csharp
using System;
using System.Windows.Markup;
using MailConvertPrivateUser.Services;

namespace MailConvertPrivateUser.Markup;

[MarkupExtensionReturnType(typeof(string))]
public class LocExtension : MarkupExtension
{
    public string Key { get; set; } = string.Empty;

    public LocExtension() { }
    public LocExtension(string key) { Key = key; }

    public override object ProvideValue(IServiceProvider serviceProvider)
    {
        if (string.IsNullOrEmpty(Key)) return string.Empty;
        return LocalizationManager.GetString(Key);
    }
}
```

- [ ] **Step 2: Build to verify compilation**

Run: `cd D:/ToolDevelop/LiteEMLTOPST/src/MailConvertPrivateUser && dotnet build 2>&1 | tail -10`
Expected: `Build succeeded.`

- [ ] **Step 3: Commit**

```bash
cd D:/ToolDevelop/LiteEMLTOPST
git add src/MailConvertPrivateUser/Markup/LocExtension.cs
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -m "feat(i18n): add LocExtension for WPF XAML binding"
```

---

## Task 6: Populate Language XML files from the catalog

**Files:**
- Modify: `Language/zh-cn.xml` (full content)
- Modify: `Language/en-us.xml` (full content)

- [ ] **Step 1: Generate the XML files from `docs/superpowers/translations/keys.csv`**

Run this bash command from the repo root to generate both files from the CSV catalog:

```bash
cd D:/ToolDevelop/LiteEMLTOPST
python -c "
import csv
for culture, fname, col in [('zh-CN', 'zh-cn.xml', 1), ('en-US', 'en-us.xml', 2)]:
    with open('docs/superpowers/translations/keys.csv', encoding='utf-8') as f, \
         open(f'Language/{fname}', 'w', encoding='utf-8') as out:
        out.write('<?xml version=\"1.0\" encoding=\"utf-8\"?>\n<resources>\n')
        for row in csv.reader(f):
            if not row or row[0].startswith('key'): continue
            out.write(f'  <string name=\"{row[0]}\">{row[col]}</string>\n')
        out.write('</resources>\n')
"
```

Verify: `wc -l Language/zh-cn.xml Language/en-us.xml` should show ~80 lines each (3 header + 2 footer + ~73 keys).

- [ ] **Step 2: Build and test to confirm the catalog is valid XML and tests still pass**

Run: `cd D:/ToolDevelop/LiteEMLTOPST && dotnet test tests/LiteEMLTOPST.Tests/LiteEMLTOPST.Tests.csproj 2>&1 | tail -5`
Expected: `Passed! - Failed: 0, Passed: 9`

- [ ] **Step 3: Commit**

```bash
cd D:/ToolDevelop/LiteEMLTOPST
git add Language/
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -m "feat(i18n): populate Language xml files from translation catalog"
```

---

## Task 7: Wire App.OnStartup to initialize LocalizationManager

**Files:**
- Modify: `src/MailConvertPrivateUser/App.xaml.cs`

- [ ] **Step 1: Add LocalizationManager initialization and replace hardcoded "错误" string**

Replace the entire content of `src/MailConvertPrivateUser/App.xaml.cs` with:

```csharp
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
```

- [ ] **Step 2: Build to verify**

Run: `cd D:/ToolDevelop/LiteEMLTOPST/src/MailConvertPrivateUser && dotnet build 2>&1 | tail -10`
Expected: `Build succeeded.`

- [ ] **Step 3: Commit**

```bash
cd D:/ToolDevelop/LiteEMLTOPST
git add src/MailConvertPrivateUser/App.xaml.cs
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -m "feat(i18n): initialize LocalizationManager from usersettings.json on startup"
```

---

## Task 8: Add Language Tab to PreferencesWindow.xaml

**Files:**
- Modify: `src/MailConvertPrivateUser/PreferencesWindow.xaml`

- [ ] **Step 1: Add `xmlns:loc` namespace to the root Window element**

In `PreferencesWindow.xaml`, change the root `<Window>` element (line 1) to:
```xml
<Window x:Class="MailConvertPrivateUser.PreferencesWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:loc="clr-namespace:MailConvertPrivateUser.Markup"
        Title="{loc:Loc Preferences_Title}" Height="540" Width="650"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        Background="#F5F7FA">
```

- [ ] **Step 2: Replace the title TextBlock in the header (line 139)**

Find: `<TextBlock Text="常规设置" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource TextPrimary}"/>`
Replace with: `<TextBlock Text="{loc:Loc Preferences_Title}" FontSize="14" FontWeight="SemiBold" Foreground="{StaticResource TextPrimary}"/>`

- [ ] **Step 3: Replace the existing 4 Tab headers with localized versions**

Find: `<TabItem Header="📧 EML转PST">`
Replace with: `<TabItem Header="{loc:Loc Preferences_Tab_EmlToPst}">`

Find: `<TabItem Header="📅 EML日期分割">`
Replace with: `<TabItem Header="{loc:Loc Preferences_Tab_EmlSplit}">`

Find: `<TabItem Header="☁️ IMAP">`
Replace with: `<TabItem Header="{loc:Loc Preferences_Tab_Imap}">`

Find: `<TabItem Header="🔄 O365">`
Replace with: `<TabItem Header="{loc:Loc Preferences_Tab_O365}">`

- [ ] **Step 4: Replace text in the EML转PST Tab**

Find: `<TextBlock Text="EML转PST 默认路径设置" FontSize="13" FontWeight="SemiBold" Foreground="{StaticResource TextPrimary}" Margin="0,0,0,12"/>`
Replace with: `<TextBlock Text="{loc:Loc Preferences_Group_EmlToPst}" FontSize="13" FontWeight="SemiBold" Foreground="{StaticResource TextPrimary}" Margin="0,0,0,12"/>`

Find: `<TextBlock Text="设置转换功能默认使用的目录路径" Foreground="{StaticResource TextSecondary}" FontSize="11" Margin="0,0,0,16"/>`
Replace with: `<TextBlock Text="{loc:Loc Preferences_Group_EmlToPstDesc}" Foreground="{StaticResource TextSecondary}" FontSize="11" Margin="0,0,0,16"/>`

Find: `<TextBlock Grid.Row="0" Grid.Column="0" Text="EML 目录" Style="{StaticResource FieldLabel}"/>`
Replace with: `<TextBlock Grid.Row="0" Grid.Column="0" Text="{loc:Loc Preferences_Field_EmlDir}" Style="{StaticResource FieldLabel}"/>`

Find: `<Button Grid.Row="0" Grid.Column="2" Content="浏览" Style="{StaticResource SecondaryButton}" Click="BrowseDefaultEmlInput_Click" Margin="8,0,0,0" Padding="12,6"/>`
Replace with: `<Button Grid.Row="0" Grid.Column="2" Content="{loc:Loc Common_Browse}" Style="{StaticResource SecondaryButton}" Click="BrowseDefaultEmlInput_Click" Margin="8,0,0,0" Padding="12,6"/>`

Find: `<TextBlock Grid.Row="1" Grid.Column="0" Text="PST 输出" Style="{StaticResource FieldLabel}" Margin="0,12,8,0"/>`
Replace with: `<TextBlock Grid.Row="1" Grid.Column="0" Text="{loc:Loc Preferences_Field_PstOutput}" Style="{StaticResource FieldLabel}" Margin="0,12,8,0"/>`

Find: `<Button Grid.Row="1" Grid.Column="2" Content="浏览" Style="{StaticResource SecondaryButton}" Click="BrowseDefaultPstOutput_Click" Margin="8,12,0,0" Padding="12,6"/>`
Replace with: `<Button Grid.Row="1" Grid.Column="2" Content="{loc:Loc Common_Browse}" Style="{StaticResource SecondaryButton}" Click="BrowseDefaultPstOutput_Click" Margin="8,12,0,0" Padding="12,6"/>`

- [ ] **Step 5: Replace text in the EML日期分割 Tab**

Find: `<TextBlock Text="EML日期分割 默认路径设置" ...`
Replace with: `<TextBlock Text="{loc:Loc Preferences_Group_EmlSplit}" ...`

Find: `<TextBlock Text="设置日期分割功能默认使用的目录路径" ...`
Replace with: `<TextBlock Text="{loc:Loc Preferences_Group_EmlSplitDesc}" ...`

Find: `<TextBlock Grid.Row="0" Grid.Column="0" Text="EML 目录" Style="{StaticResource FieldLabel}"/>`
Replace with: `<TextBlock Grid.Row="0" Grid.Column="0" Text="{loc:Loc Preferences_Field_EmlDir}" Style="{StaticResource FieldLabel}"/>`

Find: `<Button Grid.Row="0" Grid.Column="2" Content="浏览" ... Click="BrowseDefaultSplitEml_Click" .../>`
Replace with: `<Button Grid.Row="0" Grid.Column="2" Content="{loc:Loc Common_Browse}" ... Click="BrowseDefaultSplitEml_Click" .../>`

Find: `<TextBlock Grid.Row="1" Grid.Column="0" Text="日期之前" Style="{StaticResource FieldLabel}" Margin="0,12,8,0"/>`
Replace with: `<TextBlock Grid.Row="1" Grid.Column="0" Text="{loc:Loc Preferences_Field_DateBefore}" Style="{StaticResource FieldLabel}" Margin="0,12,8,0"/>`

Find: `<Button Grid.Row="1" Grid.Column="2" Content="浏览" ... Click="BrowseDefaultSplitBefore_Click" .../>`
Replace with: `<Button Grid.Row="1" Grid.Column="2" Content="{loc:Loc Common_Browse}" ... Click="BrowseDefaultSplitBefore_Click" .../>`

Find: `<TextBlock Grid.Row="2" Grid.Column="0" Text="日期之后" Style="{StaticResource FieldLabel}" Margin="0,12,8,0"/>`
Replace with: `<TextBlock Grid.Row="2" Grid.Column="0" Text="{loc:Loc Preferences_Field_DateAfter}" Style="{StaticResource FieldLabel}" Margin="0,12,8,0"/>`

Find: `<Button Grid.Row="2" Grid.Column="2" Content="浏览" ... Click="BrowseDefaultSplitAfter_Click" .../>`
Replace with: `<Button Grid.Row="2" Grid.Column="2" Content="{loc:Loc Common_Browse}" ... Click="BrowseDefaultSplitAfter_Click" .../>`

- [ ] **Step 6: Replace text in the IMAP Tab**

Find: `<TextBlock Grid.Row="0" Text="IMAP 服务器账户列表" ...`
Replace with: `<TextBlock Grid.Row="0" Text="{loc:Loc Preferences_Group_Imap}" ...`

Find the four DataGridTextColumn headers (`名称`, `服务器`, `端口`, `邮箱`) and replace with:
- `Header="{loc:Loc Common_Name}"`
- `Header="{loc:Loc Common_Server}"`
- `Header="{loc:Loc Common_Port}"`
- `Header="{loc:Loc Common_Email}"`

Find the three action buttons (`添加`, `编辑`, `删除`) and replace `Content="..."` with:
- `Content="{loc:Loc Common_Add}"`
- `Content="{loc:Loc Common_Edit}"`
- `Content="{loc:Loc Common_Delete}"`

- [ ] **Step 7: Replace text in the O365 Tab**

Find: `<TextBlock Grid.Row="0" Text="Office 365 账户列表" ...`
Replace with: `<TextBlock Grid.Row="0" Text="{loc:Loc Preferences_Group_O365}" ...`

Find: `<TextBlock Grid.Row="1" Text="配置 Azure AD 应用凭证以连接 Microsoft Graph API" ...`
Replace with: `<TextBlock Grid.Row="1" Text="{loc:Loc Preferences_Group_O365Desc}" ...`

Find the four DataGridTextColumn headers (`名称`, `Tenant ID`, `Client ID`, `用户名`) and replace with:
- `Header="{loc:Loc Common_Name}"`
- `Header="Tenant ID"` (keep as-is, technical term)
- `Header="Client ID"` (keep as-is, technical term)
- `Header="{loc:Loc Common_Username}"`

Find the three action buttons (`添加`, `编辑`, `删除`) and replace `Content="..."` with:
- `Content="{loc:Loc Common_Add}"`
- `Content="{loc:Loc Common_Edit}"`
- `Content="{loc:Loc Common_Delete}"`

- [ ] **Step 8: Replace the bottom Save/Cancel button text**

Find: `<Button Content="保存" Style="{StaticResource ModernButton}" Click="BtnSave_Click" Width="100" Margin="0,0,12,0"/>`
Replace with: `<Button Content="{loc:Loc Common_Save}" Style="{StaticResource ModernButton}" Click="BtnSave_Click" Width="100" Margin="0,0,12,0"/>`

Find: `<Button Content="取消" Style="{StaticResource SecondaryButton}" Click="BtnCancel_Click" Width="100"/>`
Replace with: `<Button Content="{loc:Loc Common_Cancel}" Style="{StaticResource SecondaryButton}" Click="BtnCancel_Click" Width="100"/>`

- [ ] **Step 9: Add the Language Tab**

Insert this new `<TabItem>` immediately before the closing `</TabControl>` tag (line 280, after the O365 Tab):

```xml
                <!-- Tab 5: Language -->
                <TabItem Header="{loc:Loc Preferences_Tab_Language}">
                    <Border Background="White" CornerRadius="0,8,8,8" Padding="16">
                        <StackPanel>
                            <TextBlock Text="{loc:Loc Preferences_Language_Prompt}" FontSize="13" FontWeight="SemiBold" Foreground="{StaticResource TextPrimary}" Margin="0,0,0,12"/>
                            <ComboBox x:Name="cmbLanguage" Width="200" HorizontalAlignment="Left" DisplayMemberPath="DisplayName" SelectedValuePath="Code"/>
                        </StackPanel>
                    </Border>
                </TabItem>
```

- [ ] **Step 10: Build to verify**

Run: `cd D:/ToolDevelop/LiteEMLTOPST/src/MailConvertPrivateUser && dotnet build 2>&1 | tail -10`
Expected: `Build succeeded.`

- [ ] **Step 11: Commit**

```bash
cd D:/ToolDevelop/LiteEMLTOPST
git add src/MailConvertPrivateUser/PreferencesWindow.xaml
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -m "feat(i18n): localize PreferencesWindow.xaml and add Language tab"
```

---

## Task 9: Wire PreferencesWindow.xaml.cs for Language save/load

**Files:**
- Modify: `src/MailConvertPrivateUser/PreferencesWindow.xaml.cs`

- [ ] **Step 1: Add the using directive and fields**

At the top of `PreferencesWindow.xaml.cs`, the existing using block:
```csharp
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using MailConvertPrivateUser.Services;
using MailConvertPrivateUser.Models;
```
is already correct. Add a new field right after `_o365Accounts`:
```csharp
    private string _selectedLanguage = "zh-CN";
```

- [ ] **Step 2: Update the constructor to populate the Language combo**

Replace the constructor with:
```csharp
    public PreferencesWindow()
    {
        InitializeComponent();
        LoadSettings();
        InitializeLanguageCombo();
    }

    private void InitializeLanguageCombo()
    {
        cmbLanguage.ItemsSource = LocalizationManager.AvailableLanguages;
        cmbLanguage.SelectedValue = LocalizationManager.CurrentLanguage;
    }
```

- [ ] **Step 3: Update `BtnSave_Click` to save the Language choice**

In `BtnSave_Click`, after the line `existingSettings["O365Accounts"] = o365Json;` and before `var jsonOut = ...`, add:

```csharp
            // 保存语言选择
            existingSettings["Language"] = (cmbLanguage.SelectedValue as string) ?? LocalizationManager.CurrentLanguage;
```

Then update the success MessageBox. Find:
```csharp
            System.Windows.MessageBox.Show("设置已保存！", "保存设置", MessageBoxButton.OK, MessageBoxImage.Information);
```
Replace with:
```csharp
            var selectedLang = (cmbLanguage.SelectedValue as string) ?? LocalizationManager.CurrentLanguage;
            var newDisplay = LocalizationManager.AvailableLanguages.FirstOrDefault(l => l.Code == selectedLang)?.DisplayName ?? selectedLang;
            // Build the message in the *current* (session) language — the new language only applies after restart.
            var restartMsg = LocalizationManager.GetString("Preferences_Language_RestartPrompt", newDisplay);
            var title = LocalizationManager.GetString("Preferences_SaveSuccessTitle");
            var successMsg = LocalizationManager.GetString("Preferences_SaveSuccess");
            System.Windows.MessageBox.Show(successMsg + "\n\n" + restartMsg, title, MessageBoxButton.OK, MessageBoxImage.Information);
```

Do NOT call `LocalizationManager.Initialize(selectedLang)` here — the current session must keep showing the old language per the spec ("Current session still shows old language, no half-translated UI"). The display name comes from the static `AvailableLanguages` list, which doesn't need a re-init to look up.

- [ ] **Step 4: Update the error MessageBox in the catch block**

Find:
```csharp
            System.Windows.MessageBox.Show($"保存设置失败: {ex.Message}", "错误", MessageBoxButton.OK, MessageBoxImage.Error);
```
Replace with:
```csharp
            System.Windows.MessageBox.Show(LocalizationManager.GetString("Preferences_SaveFailed", ex.Message), LocalizationManager.GetString("Common_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
```

- [ ] **Step 5: Update `LoadSettings` to read the Language field**

In `LoadSettings`, after the line `if (settings.TryGetValue("DefaultSplitAfter", out var defaultSplitAfter))`, add:

```csharp
                    if (settings.TryGetValue("Language", out var language) && !string.IsNullOrWhiteSpace(language))
                    {
                        _selectedLanguage = language;
                    }
```

Then, after `LoadO365AccountGrid();` (the last line of `LoadSettings`), add:
```csharp
            cmbLanguage.SelectedValue = _selectedLanguage;
```

(Move the `InitializeLanguageCombo` call's first line to set items here, OR keep it in InitializeLanguageCombo but also set selection here.)

- [ ] **Step 6: Build to verify**

Run: `cd D:/ToolDevelop/LiteEMLTOPST/src/MailConvertPrivateUser && dotnet build 2>&1 | tail -10`
Expected: `Build succeeded.`

- [ ] **Step 7: Commit**

```bash
cd D:/ToolDevelop/LiteEMLTOPST
git add src/MailConvertPrivateUser/PreferencesWindow.xaml.cs
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -m "feat(i18n): wire PreferencesWindow to save/load Language setting"
```

---

## Task 10: Localize ImapAccountEditWindow.xaml

**Files:**
- Modify: `src/MailConvertPrivateUser/ImapAccountEditWindow.xaml`

- [ ] **Step 1: Add `xmlns:loc` to the root**

Change line 1 to:
```xml
<Window x:Class="MailConvertPrivateUser.ImapAccountEditWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:loc="clr-namespace:MailConvertPrivateUser.Markup"
        Title="{loc:Loc ImapEdit_Title}" Height="340" Width="450"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        Background="#F5F7FA">
```

- [ ] **Step 2: Replace the rest of the strings**

Find `Title="编辑 IMAP 账户"` (already replaced in Step 1) ✓

Find: `<TextBlock Grid.Row="0" Text="IMAP 账户设置" ...`
Replace with: `<TextBlock Grid.Row="0" Text="{loc:Loc ImapEdit_Group}" ...`

Find: `Text="名称"` (line 54)
Replace with: `Text="{loc:Loc Common_Name}"`

Find: `Text="服务器"` (line 57)
Replace with: `Text="{loc:Loc Common_Server}"`

Find: `Text="端口"` (line 60)
Replace with: `Text="{loc:Loc Common_Port}"`

Find: `Text="邮箱"` (line 63)
Replace with: `Text="{loc:Loc Common_Email}"`

Find: `Text="密码"` (line 66)
Replace with: `Text="{loc:Loc Common_Password}"`

Find: `Text="使用 SSL"` (line 69)
Replace with: `Text="{loc:Loc ImapEdit_Field_UseSsl}"`

Find: `<Button Content="保存" Click="BtnSave_Click" ...`
Replace with: `<Button Content="{loc:Loc Common_Save}" Click="BtnSave_Click" ...`

Find: `<Button Content="取消" Click="BtnCancel_Click" ...`
Replace with: `<Button Content="{loc:Loc Common_Cancel}" Click="BtnCancel_Click" ...`

- [ ] **Step 3: Build and commit**

Run: `cd D:/ToolDevelop/LiteEMLTOPST/src/MailConvertPrivateUser && dotnet build 2>&1 | tail -5`
Expected: `Build succeeded.`

```bash
cd D:/ToolDevelop/LiteEMLTOPST
git add src/MailConvertPrivateUser/ImapAccountEditWindow.xaml
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -m "feat(i18n): localize ImapAccountEditWindow.xaml"
```

---

## Task 11: Localize O365AccountEditWindow.xaml

**Files:**
- Modify: `src/MailConvertPrivateUser/O365AccountEditWindow.xaml`

- [ ] **Step 1: Add `xmlns:loc` to the root**

Change line 1 to:
```xml
<Window x:Class="MailConvertPrivateUser.O365AccountEditWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:loc="clr-namespace:MailConvertPrivateUser.Markup"
        Title="{loc:Loc O365Edit_Title}" Height="420" Width="520"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize"
        Background="#F5F7FA">
```

- [ ] **Step 2: Replace strings**

- `Text="O365 账户设置"` → `Text="{loc:Loc O365Edit_Group}"`
- `Text="配置 Azure AD 应用凭证以连接 Microsoft Graph API"` (line 102) → `Text="{loc:Loc O365Edit_GroupDesc}"`
- `Text="配置名称"` → `Text="{loc:Loc O365Edit_Field_Name}"`
- `Text="用于标识此配置的友好名称"` → `Text="{loc:Loc O365Edit_Field_NameDesc}"`
- `Text="Tenant ID"` → keep as-is (technical)
- `Text="Azure AD 租户 ID，可在 Azure Portal - Azure Active Directory - 概述中查看"` → `Text="{loc:Loc O365Edit_Field_TenantIdDesc}"`
- `Text="Client ID"` → keep as-is
- `Text="应用程序（客户端）ID，注册 Azure AD 应用后获取"` → `Text="{loc:Loc O365Edit_Field_ClientIdDesc}"`
- `Text="用户名（可选）"` → `Text="{loc:Loc O365Edit_Field_Username}"`
- `Text="连接时使用的用户名，通常为管理员邮箱"` → `Text="{loc:Loc O365Edit_Field_UsernameDesc}"`
- `<Button Content="保存" ...` → `<Button Content="{loc:Loc Common_Save}" ...`
- `<Button Content="取消" ...` → `<Button Content="{loc:Loc Common_Cancel}" ...`

- [ ] **Step 3: Build and commit**

```bash
cd D:/ToolDevelop/LiteEMLTOPST/src/MailConvertPrivateUser && dotnet build 2>&1 | tail -5
git add src/MailConvertPrivateUser/O365AccountEditWindow.xaml
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -m "feat(i18n): localize O365AccountEditWindow.xaml"
```

---

## Task 12: Localize ContactMappingDialog.xaml

**Files:**
- Modify: `src/MailConvertPrivateUser/ContactMappingDialog.xaml`

- [ ] **Step 1: Add `xmlns:loc` to the root**

Change line 1 to:
```xml
<Window x:Class="MailConvertPrivateUser.ContactMappingDialog"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:loc="clr-namespace:MailConvertPrivateUser.Markup"
        Title="{loc:Loc ContactMapping_Title}" Height="450" Width="600"
        WindowStartupLocation="CenterOwner"
        ResizeMode="NoResize">
```

- [ ] **Step 2: Replace strings**

- `Text="请将源文件字段映射到目标字段："` → `Text="{loc:Loc ContactMapping_Header}"`
- `Content="开始转换"` → `Content="{loc:Loc ContactMapping_StartConvert}"`
- `Content="取消"` → `Content="{loc:Loc Common_Cancel}"`

- [ ] **Step 3: Build and commit**

```bash
cd D:/ToolDevelop/LiteEMLTOPST/src/MailConvertPrivateUser && dotnet build 2>&1 | tail -5
git add src/MailConvertPrivateUser/ContactMappingDialog.xaml
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -m "feat(i18n): localize ContactMappingDialog.xaml"
```

---

## Task 13: Localize RegistrationWindow.xaml

**Files:**
- Modify: `src/MailConvertPrivateUser/RegistrationWindow.xaml`

- [ ] **Step 1: Add `xmlns:loc` to the root**

Change line 1 to:
```xml
<Window x:Class="MailConvertPrivateUser.RegistrationWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:loc="clr-namespace:MailConvertPrivateUser.Markup"
        Title="{loc:Loc Reg_Title}" Height="500" Width="550"
        WindowStartupLocation="CenterScreen"
        ResizeMode="NoResize"
        Background="#F0F0F0">
```

- [ ] **Step 2: Replace all 27 strings using the catalog**

Apply these replacements (all `Text=` and `Content=` attributes, with the catalog key):

| Original (Chinese) | Key |
|---|---|
| 软件使用注册 | `Reg_Header` |
| 当前状态: | `Reg_Status_Current` |
| 试用版 | `Reg_Status_Trial` |
| 检测中... | `Reg_Status_Detecting` |
| 测试版授权信息 | `Reg_RegPanelTitle` |
| 用户姓名: | `Reg_Field_UserName` |
| 用户邮箱: | `Reg_Field_UserEmail` |
| 组织/公司: | `Reg_Field_Org` |
| 注册日期: | `Reg_Field_RegDate` |
| 到期日期: | `Reg_Field_ExpireDate` |
| MAC地址: | `Reg_Field_Mac` |
| 显示 (×2 — chkShowMac, chkShowMacInput) | `Reg_Field_ShowMac` |
| 注册（获得30天试用） | `Reg_RegFormTitle` |
| 提示：注册后可获得30天免费试用，到期后需激活订阅才能继续使用 | `Reg_RegFormHint` |
| 软件名称: | `Reg_Field_SoftwareName` |
| 版本: | `Reg_Field_Version` |
| 激活订阅版 | `Reg_ActivateTitle` |
| 提示：已有授权码？在此输入激活码升级为订阅版（无限期使用） | `Reg_ActivateHint` |
| 授权码: | `Reg_Field_Serial` |
| 激活 | `Reg_Activate` |
| 注册（试用） | `Reg_Btn_Register` |
| 注销授权 | `Reg_Btn_Unregister` |
| 关闭 | `Reg_Btn_Close` |

Example replacements (apply this pattern to each):
- `Text="软件使用注册"` → `Text="{loc:Loc Reg_Header}"`
- `Text="用户姓名:"` → `Text="{loc:Loc Reg_Field_UserName}"`

- [ ] **Step 3: Build and commit**

```bash
cd D:/ToolDevelop/LiteEMLTOPST/src/MailConvertPrivateUser && dotnet build 2>&1 | tail -5
git add src/MailConvertPrivateUser/RegistrationWindow.xaml
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -m "feat(i18n): localize RegistrationWindow.xaml"
```

---

## Task 14: Localize RegistrationWindow.xaml.cs MessageBox strings

**Files:**
- Modify: `src/MailConvertPrivateUser/RegistrationWindow.xaml.cs`

- [ ] **Step 1: Add using directive**

At the top of the file, after existing `using` statements, add:
```csharp
using MailConvertPrivateUser.Services;
```

- [ ] **Step 2: Find and replace each hardcoded MessageBox string**

Search the file for these exact Chinese strings and replace with localized calls. Use this pattern:
- `"请输入有效的邮箱地址"` → `LocalizationManager.GetString("Reg_StatusMessage_InvalidEmail")`
- `"无法获取 MAC 地址，请重试"` → `LocalizationManager.GetString("Reg_StatusMessage_MacNotFound")`
- `$"注册失败: {ex.Message}"` → `LocalizationManager.GetString("Reg_StatusMessage_RegFailed", ex.Message)`
- `"注册成功！"` → `LocalizationManager.GetString("Reg_StatusMessage_RegSuccess")`
- `"激活成功！"` → `LocalizationManager.GetString("Reg_StatusMessage_Activated")`
- `$"激活失败: {ex.Message}"` → `LocalizationManager.GetString("Reg_StatusMessage_ActivateFailed", ex.Message)`
- `"授权已注销"` → `LocalizationManager.GetString("Reg_StatusMessage_UnregSuccess")`
- `$"注销失败: {ex.Message}"` → `LocalizationManager.GetString("Reg_StatusMessage_UnregFailed", ex.Message)`
- `"确定要注销当前授权吗？"` → `LocalizationManager.GetString("Reg_StatusMessage_ConfirmUnreg")`

Also replace the `lblStatus.Text` / `lblFinalStatus.Text` Chinese assignments (e.g. `"试用版"`, `"注册成功！"`, etc.) using the same key lookup pattern.

For dynamic remaining-days text, use: `LocalizationManager.GetString("Reg_TrialRemainingDays", days)`

- [ ] **Step 3: Build and commit**

```bash
cd D:/ToolDevelop/LiteEMLTOPST/src/MailConvertPrivateUser && dotnet build 2>&1 | tail -10
git add src/MailConvertPrivateUser/RegistrationWindow.xaml.cs
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -m "feat(i18n): localize RegistrationWindow MessageBox strings"
```

---

## Task 15: Localize MainWindow.xaml

**Files:**
- Modify: `src/MailConvertPrivateUser/MainWindow.xaml`

This is the largest file (1013 lines, 226 translatable strings). It is split into sub-tasks for manageability.

- [ ] **Step 1: Add `xmlns:loc` to the root**

Change line 1 to:
```xml
<Window x:Class="MailConvertPrivateUser.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:loc="clr-namespace:MailConvertPrivateUser.Markup"
        Title="{loc:Loc Main_Title}" Height="800" Width="1420"
        WindowStartupLocation="CenterScreen"
        Background="#F5F7FA"
        Loaded="Window_Loaded">
```

- [ ] **Step 2: Replace all Tab headers and top toolbar buttons**

- `Header="📧 EML转PST"` → `Header="{loc:Loc Main_Tab_EmlToPst}"`
- `Header="📅 EML日期分割"` → `Header="{loc:Loc Main_Tab_EmlSplit}"`
- `Header="📁 目录分类"` → `Header="{loc:Loc Main_Tab_Directory}"`
- `Header="☁️ IMAP"` → `Header="{loc:Loc Main_Tab_Imap}"`
- `Header="🔄 O365"` → `Header="{loc:Loc Main_Tab_O365}"`
- `Header="👤 联系人"` → `Header="{loc:Loc Main_Tab_Contacts}"`
- `Header="🔄 IMAP转PST"` → `Header="{loc:Loc Main_Tab_ImapToPst}"`
- `Content="⚙️ 偏好设置"` → `Content="{loc:Loc Main_Btn_Preferences}"`
- `Content="🔑 注册激活"` → `Content="{loc:Loc Main_Btn_Registration}"`

- [ ] **Step 3: Replace field labels and group headers**

For every `Text="..."` attribute where the value matches a key in `docs/superpowers/translations/keys.csv` under the `Main_` prefix or common keys, replace with `{loc:Loc KeyName}`.

Use this search-and-replace approach:
1. Open `MainWindow.xaml`
2. For each `Text="<chinese>"` or `Content="<chinese>"` or `Header="<chinese>"`:
   - Look up the matching key in `docs/superpowers/translations/keys.csv`
   - Replace with `Text="{loc:Loc KeyName}"` etc.
3. Skip any text that's:
   - A `Tooltip` description that uses an English placeholder (none expected)
   - A literal punctuation/separator (e.g. ` → `, `…`)

Key references for MainWindow: `Main_Group_*`, `Main_Field_*`, `Main_Btn_*`, `Main_Tab_*`, `Main_Msg_*`, `Main_Status_*`, `Main_Table_Column_*`, plus `Common_*` for any generic terms.

- [ ] **Step 4: Build and fix any missing keys**

Run: `cd D:/ToolDevelop/LiteEMLTOPST/src/MailConvertPrivateUser && dotnet build 2>&1 | tail -20`
Expected: `Build succeeded.`

If any key was missed, the build still passes (LocExtension returns `[key]`), but a string like `[Main_Field_X]` will appear in the UI at runtime. Note any untranslated keys for Task 18.

- [ ] **Step 5: Commit**

```bash
cd D:/ToolDevelop/LiteEMLTOPST
git add src/MailConvertPrivateUser/MainWindow.xaml
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -m "feat(i18n): localize MainWindow.xaml"
```

---

## Task 16: Localize MainWindow.xaml.cs MessageBox strings

**Files:**
- Modify: `src/MailConvertPrivateUser/MainWindow.xaml.cs`

- [ ] **Step 1: Add using directive**

At the top of the file (after existing `using` statements), add:
```csharp
using MailConvertPrivateUser.Services;
```

- [ ] **Step 2: Replace all hardcoded Chinese strings in MessageBox.Show calls**

Search the file for `MessageBox.Show(` and within those calls, replace any Chinese string literal with the corresponding `LocalizationManager.GetString("Key", ...)` call.

Key references:
- `"确定要退出吗？"` → `LocalizationManager.GetString("Main_Msg_ConfirmExit")`
- `"转换完成！共处理 {0} 封邮件"` → `LocalizationManager.GetString("Main_Msg_ConvertComplete", count)`
- `"转换失败: {0}"` → `LocalizationManager.GetString("Main_Msg_ConvertFailed", ex.Message)`
- `"未找到 EML 文件"` → `LocalizationManager.GetString("Main_Msg_NoEmlFiles")`
- `"请选择有效的目录"` → `LocalizationManager.GetString("Main_Msg_InvalidPath")`
- `"PST 写入失败: {0}"` → `LocalizationManager.GetString("Main_Msg_PstWriteFailed", ex.Message)`
- `"确定要删除吗？"` → `LocalizationManager.GetString("Main_Msg_ConfirmDelete")`
- `"IMAP 连接失败: {0}"` → `LocalizationManager.GetString("Main_Msg_ImapConnectFailed", ex.Message)`
- `"O365 连接失败: {0}"` → `LocalizationManager.GetString("Main_Msg_O365ConnectFailed", ex.Message)`

Also replace any `Title="..."` in `new FolderBrowserDialog()` or MessageBox titles that are Chinese with the appropriate Common_/Main_ key.

For any Chinese text assigned to a status bar / log line (`txtStatus.Text = "..."`), use the Main_Status_ keys.

- [ ] **Step 3: Build and commit**

```bash
cd D:/ToolDevelop/LiteEMLTOPST/src/MailConvert/LiteEMLTOPST/src/MailConvertPrivateUser && dotnet build 2>&1 | tail -10
git add src/MailConvertPrivateUser/MainWindow.xaml.cs
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -m "feat(i18n): localize MainWindow MessageBox and status strings"
```

---

## Task 17: Final build verification

**Files:** (none modified)

- [ ] **Step 1: Build the entire solution**

Run: `cd D:/ToolDevelop/LiteEMLTOPST && dotnet build 2>&1 | tail -15`
Expected: `Build succeeded.` with 0 errors. Warnings are OK.

- [ ] **Step 2: Run all tests**

Run: `cd D:/ToolDevelop/LiteEMLTOPST && dotnet test 2>&1 | tail -10`
Expected: `Passed! - Failed: 0, Passed: 9`

- [ ] **Step 3: Smoke test the app**

Run: `cd D:/ToolDevelop/LiteEMLTOPST && powershell -Command "Start-Process 'src\MailConvertPrivateUser\bin\Debug\net8.0-windows\MailConvertPrivateUser.exe'"`
Then verify in the UI:
- [ ] App launches in Chinese by default
- [ ] Open 偏好设置 (Preferences) → see 5 tabs including 🌐 语言
- [ ] Switch to English → Save → see "Language switched to English. Please restart..." prompt
- [ ] Restart app → UI now in English
- [ ] Switch back to 中文 → Save → restart → UI in Chinese
- [ ] Open and close Registration window — text is translated
- [ ] Open Preferences → IMAP → Add → text is translated

---

## Task 18: Catalog gap check and version bump

**Files:**
- Modify: `docs/superpowers/translations/keys.csv` (add any missing keys)
- Modify: `Language/zh-cn.xml`, `Language/en-us.xml` (sync from CSV)
- Modify: `README.md` (note new language feature)
- Modify: `src/MailConvertPrivateUser/MailConvertPrivateUser.csproj` (version bump — if there's a `<Version>` field; otherwise add to assembly info or skip)

- [ ] **Step 1: Hunt for untranslated Chinese strings in the source**

Run: `cd D:/ToolDevelop/LiteEMLTOPST && grep -rn '[\xe4-\xe9][\x80-\xbf][\x80-\xbf]' src/MailConvertPrivateUser/*.xaml src/MailConvertPrivateUser/*.cs | grep -v 'Language/\|bin/\|obj/'`
For each match that is user-facing (not a log message, not a path, not a regex), add it to `docs/superpowers/translations/keys.csv` with a fresh key, then regenerate both XML files using the python script from Task 6.

- [ ] **Step 2: Rebuild and re-test**

Run: `cd D:/ToolDevelop/LiteEMLTOPST && dotnet build 2>&1 | tail -5 && dotnet test 2>&1 | tail -5`
Expected: build succeeds, all 9 tests pass.

- [ ] **Step 3: Update README to document the feature**

Add a short "## 语言切换 / Language Switching" section to `README.md` (near the existing 功能特性 list):

```markdown
## 语言切换 / Language Switching

支持中文和英文。打开 **偏好设置 → 🌐 语言**，选择语言后保存并重启应用即可。

Supports Chinese and English. Open **Preferences → 🌐 Language**, select a language, save, and restart the app to apply.
```

- [ ] **Step 4: Commit**

```bash
cd D:/ToolDevelop/LiteEMLTOPST
git add Language/ docs/superpowers/translations/ README.md src/MailConvertPrivateUser/MailConvertPrivateUser.csproj
git -c user.name="Claude" -c user.email="noreply@anthropic.com" commit -m "feat(i18n): fill catalog gaps, update README, bump version to v1.0.0.7"
```

---

## Self-Review Notes

**Spec coverage check:**
- ✅ XML files (`zh-cn.xml`, `en-us.xml`) — Tasks 1, 6
- ✅ `LocalizationManager` service — Tasks 3, 4
- ✅ `LocExtension` — Task 5
- ✅ `usersettings.json["Language"]` read/write — Tasks 7, 9
- ✅ Restart prompt — Task 9
- ✅ All 6 XAML windows refactored — Tasks 8, 10, 11, 12, 13, 15
- ✅ Code-behind MessageBox strings — Tasks 7, 14, 16
- ✅ `[key]` bracketed fallback — Task 4 (GetString implementation)
- ✅ zh-cn fallback when current culture key missing — Task 4 (LoadInto + _fallback)
- ✅ Hardcoded error in App.xaml.cs — Task 7
- ✅ Version bump — Task 18
- ✅ Test coverage — Task 3 (xunit tests for LocalizationManager)

**Placeholder scan:** No TBDs, TODOs, or "implement later" markers. All steps include concrete code or file paths.

**Type consistency:**
- `LocalizationManager.Initialize(string culture)` — used identically in Tasks 4, 7, 9
- `LocalizationManager.GetString(string key, params object[] args)` — used in Tasks 7, 9, 14, 16
- `LocalizationManager.AvailableLanguages` — used in Tasks 4, 9
- `LocExtension.Key` property — used in all XAML tasks via `{loc:Loc KeyName}`
- `LanguageInfo` record — used in Task 4 (defined) and Task 9 (consumed)

**Open issues:** None.
