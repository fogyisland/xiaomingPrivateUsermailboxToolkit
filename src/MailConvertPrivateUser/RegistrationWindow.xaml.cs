using System;
using System.Net;
using System.Net.Mail;
using System.Windows;
using MailConvertPrivateUser.Services;
using Serilog;

namespace MailConvertPrivateUser;

public partial class RegistrationWindow : Window
{
    private readonly RegistrationService _registrationService;
    private string _macAddress = "";
    private bool _isRegistered;
    private readonly string _softwareName = "xiaomingMailToolkitPrivate";
    private readonly string _softwareVersion = "1.2.0.0";

    public RegistrationWindow()
    {
        InitializeComponent();
        _registrationService = new RegistrationService();

        var settings = RegistryService.LoadRegistrationInfo();
        _isRegistered = settings.IsRegistered;

        LoadMacAddress();
        UpdateUI(settings);
    }

    private void UpdateUI(AppSettings settings)
    {
        if (_isRegistered)
        {
            pnlTrial.Visibility = Visibility.Collapsed;
            pnlRegistered.Visibility = Visibility.Visible;
            pnlRegistration.Visibility = Visibility.Collapsed;
            btnRegister.Visibility = Visibility.Collapsed;
            btnUnregister.Visibility = Visibility.Visible;

            lblUserName.Text = settings.RegisteredUserName;
            lblUserEmail.Text = settings.RegisteredUserEmail;
            lblOrg.Text = string.IsNullOrEmpty(settings.RegisteredOrganization) ? "-" : settings.RegisteredOrganization;
            lblRegDate.Text = settings.RegisterDate?.ToString("yyyy-MM-dd") ?? "-";
            lblExpireDate.Text = string.IsNullOrEmpty(settings.RegisterExpireDate) ? "-" : settings.RegisterExpireDate;
            lblMacAddress.Text = MaskMac(settings.RegisteredMacAddress);

            lblStatus.Text = string.IsNullOrEmpty(settings.RegisterSerialNumber) ? LocalizationManager.GetString("Reg_Status_Trial") : LocalizationManager.GetString("Reg_Status_Registered");
            lblStatus.Foreground = string.IsNullOrEmpty(settings.RegisterSerialNumber) ?
                new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(255, 152, 0)) :
                new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(76, 175, 80));

            if (settings.RegisterRemainingDays.HasValue)
            {
                lblRemainingDays.Text = LocalizationManager.GetString("Reg_TrialRemainingDays", settings.RegisterRemainingDays.Value);
            }

            txtSerialNumber.Text = settings.RegisterSerialNumber ?? "";
        }
        else
        {
            pnlTrial.Visibility = Visibility.Visible;
            pnlRegistered.Visibility = Visibility.Collapsed;
            pnlRegistration.Visibility = Visibility.Visible;
            btnRegister.Visibility = Visibility.Visible;
            btnUnregister.Visibility = Visibility.Collapsed;

            lblStatus.Text = LocalizationManager.GetString("Reg_Status_Trial");
            lblRemainingDays.Text = LocalizationManager.GetString("Reg_TrialRemainingDays", 30);
        }
    }

    private void LoadMacAddress()
    {
        try
        {
            _macAddress = _registrationService.GetPhysicalMacAddress();

            if (string.IsNullOrEmpty(_macAddress))
            {
                lblMacAddressInput.Text = LocalizationManager.GetString("Reg_Status_NoNetworkAdapter");
                lblMacAddressInput.Foreground = System.Windows.Media.Brushes.Red;
            }
            else
            {
                lblMacAddressInput.Text = MaskMac(_macAddress);
            }

            // 更新已注册界面的MAC地址
            var settings = RegistryService.LoadRegistrationInfo();
            if (!string.IsNullOrEmpty(settings.RegisteredMacAddress))
            {
                lblMacAddress.Text = MaskMac(settings.RegisteredMacAddress);
            }
        }
        catch (Exception ex)
        {
            Log.Error(ex, "加载MAC地址失败");
            lblMacAddressInput.Text = LocalizationManager.GetString("Reg_Status_DetectFailed");
        }
    }

    private string MaskMac(string mac)
    {
        if (string.IsNullOrEmpty(mac) || mac.Length < 8)
            return mac;
        return mac.Substring(0, 2) + "-**-**-**-**-" + mac.Substring(mac.Length - 2);
    }

    private string UnmaskMac(string mac)
    {
        // 如果是掩码，返回存储的完整MAC
        var settings = RegistryService.LoadRegistrationInfo();
        return settings.RegisteredMacAddress;
    }

    private void chkShowMac_Click(object sender, RoutedEventArgs e)
    {
        var settings = RegistryService.LoadRegistrationInfo();
        lblMacAddress.Text = chkShowMac.IsChecked == true ?
            settings.RegisteredMacAddress : MaskMac(settings.RegisteredMacAddress);
    }

    private void chkShowMacInput_Click(object sender, RoutedEventArgs e)
    {
        lblMacAddressInput.Text = chkShowMacInput.IsChecked == true ?
            _macAddress : MaskMac(_macAddress);
    }

    private bool IsValidEmail(string email)
    {
        try
        {
            var addr = new MailAddress(email);
            return addr.Address == email;
        }
        catch
        {
            return false;
        }
    }

    private async void BtnRegister_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtUserName.Text))
        {
            lblFinalStatus.Text = LocalizationManager.GetString("Reg_Validation_NoUserName");
            lblFinalStatus.Foreground = System.Windows.Media.Brushes.Red;
            return;
        }

        if (string.IsNullOrWhiteSpace(txtUserEmail.Text))
        {
            lblFinalStatus.Text = LocalizationManager.GetString("Reg_Validation_NoUserEmail");
            lblFinalStatus.Foreground = System.Windows.Media.Brushes.Red;
            return;
        }

        if (!IsValidEmail(txtUserEmail.Text))
        {
            lblFinalStatus.Text = LocalizationManager.GetString("Reg_StatusMessage_InvalidEmail");
            lblFinalStatus.Foreground = System.Windows.Media.Brushes.Red;
            return;
        }

        if (string.IsNullOrEmpty(_macAddress))
        {
            lblFinalStatus.Text = LocalizationManager.GetString("Reg_Validation_NoNetworkForRegister");
            lblFinalStatus.Foreground = System.Windows.Media.Brushes.Red;
            return;
        }

        btnRegister.IsEnabled = false;
        lblFinalStatus.Text = LocalizationManager.GetString("Reg_Status_Submitting");
        lblFinalStatus.Foreground = System.Windows.Media.Brushes.Blue;

        try
        {
            var result = await _registrationService.RegisterAsync(
                _softwareName,
                _softwareVersion,
                txtUserName.Text.Trim(),
                txtUserEmail.Text.Trim(),
                txtOrganization.Text.Trim(),
                _macAddress
            );

            if (result.Success)
            {
                var settings = new AppSettings();
                settings.IsRegistered = true;
                settings.RegisteredUserName = txtUserName.Text.Trim();
                settings.RegisteredUserEmail = txtUserEmail.Text.Trim();
                settings.RegisteredOrganization = txtOrganization.Text.Trim();
                settings.RegisteredMacAddress = _macAddress;
                settings.RegisterSerialNumber = "";
                settings.RegisterDate = DateTime.Now;
                settings.RegisterExpireDate = result.ExpireDate;

                if (DateTime.TryParse(result.ExpireDate, out var expireDate))
                {
                    settings.RegisterRemainingDays = (int)Math.Max(0, (expireDate - DateTime.Now).TotalDays);
                }
                else
                {
                    settings.RegisterRemainingDays = result.RemainingDays;
                }

                // 只保存到注册表
                RegistryService.SaveRegistrationInfo(settings);

                lblFinalStatus.Text = LocalizationManager.GetString("Reg_Status_RegSuccessWithDays", settings.RegisterRemainingDays);
                lblFinalStatus.Foreground = System.Windows.Media.Brushes.Green;

                await System.Threading.Tasks.Task.Delay(2000);
                _isRegistered = true;
                UpdateUI(settings);

                if (Owner is MainWindow mainWindow)
                {
                    mainWindow.UpdateRegistrationStatus();
                }
            }
            else
            {
                lblFinalStatus.Text = result.Message;
                lblFinalStatus.Foreground = System.Windows.Media.Brushes.Red;
                btnRegister.IsEnabled = true;
            }
        }
        catch (Exception ex)
        {
            Log.Error(ex, "注册异常");
            lblFinalStatus.Text = LocalizationManager.GetString("Reg_Status_RegException", ex.Message);
            lblFinalStatus.Foreground = System.Windows.Media.Brushes.Red;
            btnRegister.IsEnabled = true;
        }
    }

    private async void BtnActivate_Click(object sender, RoutedEventArgs e)
    {
        if (string.IsNullOrWhiteSpace(txtSerialNumber.Text))
        {
            lblStatusMessage.Text = LocalizationManager.GetString("Reg_Validation_NoSerial");
            lblStatusMessage.Foreground = System.Windows.Media.Brushes.Red;
            return;
        }

        if (string.IsNullOrEmpty(_macAddress))
        {
            lblStatusMessage.Text = LocalizationManager.GetString("Reg_Validation_NoNetworkForActivate");
            lblStatusMessage.Foreground = System.Windows.Media.Brushes.Red;
            return;
        }

        btnActivate.IsEnabled = false;
        lblStatusMessage.Text = LocalizationManager.GetString("Reg_Status_Activating");
        lblStatusMessage.Foreground = System.Windows.Media.Brushes.Blue;

        try
        {
            var settings = RegistryService.LoadRegistrationInfo();
            Log.Information("[激活] 开始激活: SerialNumber={Serial}, Mac={Mac}, User={User}, Email={Email}",
                txtSerialNumber.Text.Trim(), settings.RegisteredMacAddress, settings.RegisteredUserName, settings.RegisteredUserEmail);

            var result = await _registrationService.ActivateByCodeAsync(
                txtSerialNumber.Text.Trim(),
                settings.RegisteredMacAddress,
                settings.RegisteredUserName,
                settings.RegisteredUserEmail,
                DateTime.Now.ToString("yyyy-MM-dd")
            );

            Log.Information("[激活] 结果: Success={Success}, Message={Message}", result.Success, result.Message);

            if (result.Success)
            {
                settings.RegisterSerialNumber = txtSerialNumber.Text.Trim();

                DateTime currentExpireDate;
                bool hasCurrentSerial = !string.IsNullOrEmpty(settings.RegisterSerialNumber);
                bool isExpired = !DateTime.TryParse(settings.RegisterExpireDate, out currentExpireDate) || currentExpireDate <= DateTime.Now;

                if (!isExpired && result.TotalDays.HasValue)
                {
                    var newExpireDate = currentExpireDate.AddDays(result.TotalDays.Value);
                    settings.RegisterExpireDate = newExpireDate.ToString("yyyy-MM-dd");
                    settings.RegisterRemainingDays = (int)Math.Max(0, (newExpireDate - DateTime.Now).TotalDays);
                }
                else
                {
                    var activateDate = DateTime.Now;
                    if (result.TotalDays.HasValue)
                    {
                        var newExpireDate = activateDate.AddDays(result.TotalDays.Value);
                        settings.RegisterExpireDate = newExpireDate.ToString("yyyy-MM-dd");
                        settings.RegisterRemainingDays = result.TotalDays.Value;
                        settings.RegisterDate = activateDate;
                    }
                    else
                    {
                        settings.RegisterExpireDate = result.ExpireDate;
                        settings.RegisterRemainingDays = result.RemainingDays;
                    }
                }

                if (!settings.FirstRunDate.HasValue)
                {
                    settings.FirstRunDate = DateTime.Now;
                }

                // 只保存到注册表
                RegistryService.SaveRegistrationInfo(settings);

                lblStatusMessage.Text = LocalizationManager.GetString("Reg_Status_ActivateSuccessWithDays", settings.RegisterRemainingDays);
                lblStatusMessage.Foreground = System.Windows.Media.Brushes.Green;

                await System.Threading.Tasks.Task.Delay(2000);
                UpdateUI(settings);

                if (Owner is MainWindow mainWindow)
                {
                    mainWindow.UpdateRegistrationStatus();
                }
            }
            else
            {
                lblStatusMessage.Text = result.Message;
                lblStatusMessage.Foreground = System.Windows.Media.Brushes.Red;
                btnActivate.IsEnabled = true;
            }
        }
        catch (Exception ex)
        {
            Log.Error(ex, "激活异常");
            lblStatusMessage.Text = LocalizationManager.GetString("Reg_Status_ActivateException", ex.Message);
            lblStatusMessage.Foreground = System.Windows.Media.Brushes.Red;
            btnActivate.IsEnabled = true;
        }
    }

    private void BtnUnregister_Click(object sender, RoutedEventArgs e)
    {
        var confirmBody = LocalizationManager.GetString("Reg_StatusMessage_ConfirmUnreg")
            + "\n"
            + LocalizationManager.GetString("Reg_StatusMessage_UnregAfterNote");
        var result = System.Windows.MessageBox.Show(
            confirmBody,
            LocalizationManager.GetString("Reg_Title_ConfirmUnreg"),
            MessageBoxButton.YesNo,
            MessageBoxImage.Warning);

        if (result == MessageBoxResult.Yes)
        {
            SettingsService.Clear();
            RegistryService.ClearRegistrationInfo();

            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Reg_StatusMessage_UnregSuccessNote"),
                LocalizationManager.GetString("Common_Info"),
                MessageBoxButton.OK,
                MessageBoxImage.Information);

            _isRegistered = false;
            var settings = new AppSettings();
            UpdateUI(settings);

            if (Owner is MainWindow mainWindow)
            {
                mainWindow.UpdateRegistrationStatus();
            }
        }
    }

    private void BtnClose_Click(object sender, RoutedEventArgs e)
    {
        Close();
    }
}
