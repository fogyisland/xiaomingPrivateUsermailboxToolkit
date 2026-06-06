using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Windows;
using MailConvertPrivateUser.Services;
using MailConvertPrivateUser.Models;

namespace MailConvertPrivateUser;

public partial class PreferencesWindow : Window
{
    private List<ImapAccountConfig> _imapAccounts = new();
    private List<O365AccountConfig> _o365Accounts = new();
    private string _selectedLanguage = "zh-CN";

    public PreferencesWindow()
    {
        InitializeComponent();
        InitializeLanguageCombo();
        LoadSettings();
    }

    private void InitializeLanguageCombo()
    {
        cmbLanguage.ItemsSource = LocalizationManager.AvailableLanguages;
        cmbLanguage.SelectedValue = LocalizationManager.CurrentLanguage;
    }

    private void BrowseDefaultEmlInput_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            txtDefaultEmlInput.Text = dialog.SelectedPath;
        }
    }

    private void BrowseDefaultPstOutput_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            txtDefaultPstOutput.Text = dialog.SelectedPath;
        }
    }

    private void BrowseDefaultSplitEml_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            txtDefaultSplitEml.Text = dialog.SelectedPath;
        }
    }

    private void BrowseDefaultSplitBefore_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            txtDefaultSplitBefore.Text = dialog.SelectedPath;
        }
    }

    private void BrowseDefaultSplitAfter_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            txtDefaultSplitAfter.Text = dialog.SelectedPath;
        }
    }

    private void AddImapAccount_Click(object sender, RoutedEventArgs e)
    {
        var editWindow = new ImapAccountEditWindow(new ImapAccountConfig { Port = "993", UseSsl = true });
        if (editWindow.ShowDialog() == true)
        {
            _imapAccounts.Add(editWindow.Account);
            LoadImapAccountGrid();
        }
    }

    private void EditImapAccount_Click(object sender, RoutedEventArgs e)
    {
        if (dgImapAccounts.SelectedItem is ImapAccountConfig selected)
        {
            var editWindow = new ImapAccountEditWindow(selected);
            if (editWindow.ShowDialog() == true)
            {
                // 更新列表中的项
                var index = _imapAccounts.FindIndex(a => a.Email == selected.Email);
                if (index >= 0)
                {
                    _imapAccounts[index] = editWindow.Account;
                    LoadImapAccountGrid();
                }
            }
        }
    }

    private void DeleteImapAccount_Click(object sender, RoutedEventArgs e)
    {
        if (dgImapAccounts.SelectedItem is ImapAccountConfig selected)
        {
            _imapAccounts.Remove(selected);
            LoadImapAccountGrid();
        }
    }

    private void LoadImapAccountGrid()
    {
        dgImapAccounts.ItemsSource = null;
        dgImapAccounts.ItemsSource = _imapAccounts;
    }

    private void BtnSave_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var settings = new Dictionary<string, string>
            {
                { "DefaultEmlInput", txtDefaultEmlInput.Text },
                { "DefaultPstOutput", txtDefaultPstOutput.Text },
                { "DefaultSplitEml", txtDefaultSplitEml.Text },
                { "DefaultSplitBefore", txtDefaultSplitBefore.Text },
                { "DefaultSplitAfter", txtDefaultSplitAfter.Text }
            };

            var appSettingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "usersettings.json");
            var existingSettings = new Dictionary<string, string>();

            if (File.Exists(appSettingsPath))
            {
                var json = File.ReadAllText(appSettingsPath);
                var existing = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                if (existing != null)
                {
                    existingSettings = existing;
                }
            }

            foreach (var kvp in settings)
            {
                existingSettings[kvp.Key] = kvp.Value;
            }

            // 保存 IMAP 账户列表
            var imapJson = System.Text.Json.JsonSerializer.Serialize(_imapAccounts, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
            existingSettings["ImapAccounts"] = imapJson;

            // 保存 O365 账户列表
            var o365Json = System.Text.Json.JsonSerializer.Serialize(_o365Accounts, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
            existingSettings["O365Accounts"] = o365Json;

            // 保存语言选择
            existingSettings["Language"] = (cmbLanguage.SelectedValue as string) ?? LocalizationManager.CurrentLanguage;

            var jsonOut = System.Text.Json.JsonSerializer.Serialize(existingSettings, new System.Text.Json.JsonSerializerOptions { WriteIndented = true });
            File.WriteAllText(appSettingsPath, jsonOut);

            // 提示用户新语言将在重启后生效
            var selectedLang = (cmbLanguage.SelectedValue as string) ?? LocalizationManager.CurrentLanguage;
            var newDisplay = LocalizationManager.AvailableLanguages.FirstOrDefault(l => l.Code == selectedLang)?.DisplayName ?? selectedLang;
            // 当前会话仍使用旧语言；新语言在 App.OnStartup 重启后由 LocalizationManager.Initialize 加载
            var restartMsg = LocalizationManager.GetString("Preferences_Language_RestartPrompt", newDisplay);
            var title = LocalizationManager.GetString("Preferences_SaveSuccessTitle");
            var successMsg = LocalizationManager.GetString("Preferences_SaveSuccess");
            System.Windows.MessageBox.Show(successMsg + "\n\n" + restartMsg, title, MessageBoxButton.OK, MessageBoxImage.Information);
            DialogResult = true;
            Close();
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(LocalizationManager.GetString("Preferences_SaveFailed", ex.Message), LocalizationManager.GetString("Common_Error"), MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BtnCancel_Click(object sender, RoutedEventArgs e)
    {
        DialogResult = false;
        Close();
    }

    private void LoadSettings()
    {
        try
        {
            var appSettingsPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "usersettings.json");
            if (File.Exists(appSettingsPath))
            {
                var json = File.ReadAllText(appSettingsPath);
                var settings = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(json);

                if (settings != null)
                {
                    if (settings.TryGetValue("DefaultEmlInput", out var defaultEmlInput))
                        txtDefaultEmlInput.Text = defaultEmlInput;
                    if (settings.TryGetValue("DefaultPstOutput", out var defaultPstOutput))
                        txtDefaultPstOutput.Text = defaultPstOutput;
                    if (settings.TryGetValue("DefaultSplitEml", out var defaultSplitEml))
                        txtDefaultSplitEml.Text = defaultSplitEml;
                    if (settings.TryGetValue("DefaultSplitBefore", out var defaultSplitBefore))
                        txtDefaultSplitBefore.Text = defaultSplitBefore;
                    if (settings.TryGetValue("DefaultSplitAfter", out var defaultSplitAfter))
                        txtDefaultSplitAfter.Text = defaultSplitAfter;

                    if (settings.TryGetValue("Language", out var language) && !string.IsNullOrWhiteSpace(language))
                    {
                        _selectedLanguage = language;
                    }

                    // 加载 IMAP 账户列表
                    if (settings.TryGetValue("ImapAccounts", out var imapAccountsJson) && !string.IsNullOrEmpty(imapAccountsJson))
                    {
                        var accounts = System.Text.Json.JsonSerializer.Deserialize<List<ImapAccountConfig>>(imapAccountsJson);
                        if (accounts != null)
                        {
                            _imapAccounts = accounts;
                        }
                    }

                    // 加载 O365 账户列表
                    if (settings.TryGetValue("O365Accounts", out var o365AccountsJson) && !string.IsNullOrEmpty(o365AccountsJson))
                    {
                        var accounts = System.Text.Json.JsonSerializer.Deserialize<List<O365AccountConfig>>(o365AccountsJson);
                        if (accounts != null)
                        {
                            _o365Accounts = accounts;
                        }
                    }
                }
            }

            LoadImapAccountGrid();
            LoadO365AccountGrid();
            cmbLanguage.SelectedValue = _selectedLanguage;
        }
        catch { }
    }

    #region O365 Account Management

    private void AddO365Account_Click(object sender, RoutedEventArgs e)
    {
        var newAccount = new O365AccountConfig { Name = "新配置", TenantId = "", ClientId = "", Username = "" };
        var editWindow = new O365AccountEditWindow(newAccount) { Owner = this };
        if (editWindow.ShowDialog() == true)
        {
            _o365Accounts.Add(editWindow.Account);
            LoadO365AccountGrid();
        }
    }

    private void EditO365Account_Click(object sender, RoutedEventArgs e)
    {
        if (dgO365Accounts.SelectedItem is O365AccountConfig selected)
        {
            var editWindow = new O365AccountEditWindow(selected) { Owner = this };
            if (editWindow.ShowDialog() == true)
            {
                LoadO365AccountGrid();
            }
        }
    }

    private void DeleteO365Account_Click(object sender, RoutedEventArgs e)
    {
        if (dgO365Accounts.SelectedItem is O365AccountConfig selected)
        {
            _o365Accounts.Remove(selected);
            LoadO365AccountGrid();
        }
    }

    private void LoadO365AccountGrid()
    {
        dgO365Accounts.ItemsSource = null;
        dgO365Accounts.ItemsSource = _o365Accounts;
    }

    #endregion
}
