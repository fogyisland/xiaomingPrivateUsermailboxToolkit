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

    [Fact]
    public void GetString_ReturnsEmptyString_WhenKeyIsNullOrEmpty()
    {
        Assert.Equal(string.Empty, LocalizationManager.GetString(null!));
        Assert.Equal(string.Empty, LocalizationManager.GetString(string.Empty));
    }

    [Fact]
    public void GetString_SupportsMultipleFormatPlaceholders()
    {
        WriteXml("zh-cn.xml", "<?xml version=\"1.0\"?><resources><string name=\"WelcomeMulti\">欢迎 {0}，你有 {1} 封邮件</string></resources>");
        WriteXml("en-us.xml", "<?xml version=\"1.0\"?><resources><string name=\"WelcomeMulti\">Welcome {0}, you have {1} messages</string></resources>");

        LocalizationManager.Initialize("en-US");

        Assert.Equal("Welcome Alice, you have 3 messages", LocalizationManager.GetString("WelcomeMulti", "Alice", 3));
    }

    [Fact]
    public void GetString_ReturnsTemplateVerbatim_WhenNoArgsProvidedOnTemplatedKey()
    {
        WriteXml("zh-cn.xml", "<?xml version=\"1.0\"?><resources><string name=\"Welcome\">欢迎 {0}</string></resources>");
        WriteXml("en-us.xml", "<?xml version=\"1.0\"?><resources><string name=\"Welcome\">Welcome {0}</string></resources>");

        LocalizationManager.Initialize("en-US");

        // No args passed — should return template as-is, not throw
        Assert.Equal("Welcome {0}", LocalizationManager.GetString("Welcome"));
    }

    [Fact]
    public void Initialize_FallsBackToZhCN_WhenCultureIsUnsupported()
    {
        WriteXml("zh-cn.xml", "<?xml version=\"1.0\"?><resources><string name=\"Greeting\">你好</string></resources>");

        LocalizationManager.Initialize("fr-FR");
        Assert.Equal("zh-CN", LocalizationManager.CurrentLanguage);
        Assert.Equal("你好", LocalizationManager.GetString("Greeting"));
    }
}
