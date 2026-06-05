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
