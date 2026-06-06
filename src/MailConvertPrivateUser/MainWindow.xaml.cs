using System.Windows;
using System.Runtime.InteropServices;
using Microsoft.Win32;
using MailConvertPrivateUser.Services;
using MailConvertPrivateUser.Models;
using Serilog;
using MimeKit;
using MimeKit.Text;

namespace MailConvertPrivateUser;

/// <summary>
/// OST 账户信息
/// </summary>
public class OstStoreInfo
{
    public string DisplayName { get; set; } = "";
    public string FilePath { get; set; } = "";
}

public partial class MainWindow : Window
{
    private readonly ImapExtractService _imapService = new();
    private readonly PstExtractService _pstExtractService = new();
    private readonly OstExtractService _ostExtractService = new();
    private Office365SyncService _o365Service = new();
    private readonly object _logLock = new();
    private string? _selectedPstFilePath;
    private string? _selectedEmlFolderPath;

    public MainWindow()
    {
        InitializeComponent();
    }

    private void Window_Loaded(object sender, RoutedEventArgs e)
    {
        // 检查注册状态
        if (!CheckRegistration())
        {
            // 未注册，直接退出程序
            Environment.Exit(0);
            return;
        }

        AppendLog("主窗口初始化完成");
        LoadSettings();
        LoadImapConfigList();
        LoadO365ConfigList();
        UpdateVersionLabel();

        // 启动时发送待发送的遥测数据
        _ = SendPendingTelemetryAsync();
    }

    private bool CheckRegistration()
    {
        try
        {
            var regSettings = RegistryService.LoadRegistrationInfo();

            // 未注册
            if (!regSettings.IsRegistered)
            {
                var regWindow = new RegistrationWindow { Owner = this };
                if (regWindow.ShowDialog() != true)
                {
                    // 用户取消注册，退出程序
                    return false;
                }
                regSettings = RegistryService.LoadRegistrationInfo();
                return regSettings.IsRegistered;
            }

            // 已过期
            if (!string.IsNullOrEmpty(regSettings.RegisterExpireDate) &&
                DateTime.TryParse(regSettings.RegisterExpireDate, out var expireDate) &&
                expireDate <= DateTime.Now)
            {
                var regWindow = new RegistrationWindow { Owner = this };
                if (regWindow.ShowDialog() != true)
                {
                    // 用户取消注册，退出程序
                    return false;
                }
                regSettings = RegistryService.LoadRegistrationInfo();
                return regSettings.IsRegistered;
            }

            return true;
        }
        catch (Exception ex)
        {
            Log.Error(ex, "[MainWindow] Registration check failed");
            return false;
        }
    }

    private void UpdateVersionLabel()
    {
        var settings = RegistryService.LoadRegistrationInfo();

        // 动态计算剩余天数，基于过期日期计算
        int? remainingDays = null;
        if (!string.IsNullOrEmpty(settings.RegisterExpireDate) &&
            DateTime.TryParse(settings.RegisterExpireDate, out var expireDate))
        {
            remainingDays = (int)Math.Max(0, (expireDate - DateTime.Now).TotalDays);
        }

        if (settings.IsRegistered)
        {
            // 有序列号才显示订阅版，没有序列号则是试用版
            bool isActivated = !string.IsNullOrEmpty(settings.RegisterSerialNumber);
            if (isActivated)
            {
                if (remainingDays.HasValue && remainingDays.Value > 0)
                {
                    lblVersion.Text = LocalizationManager.GetString("Main_Lbl_SubscriptionRemaining", remainingDays.Value);
                }
                else
                {
                    lblVersion.Text = LocalizationManager.GetString("Main_Lbl_Subscription");
                }
            }
            else
            {
                if (remainingDays.HasValue && remainingDays.Value > 0)
                {
                    lblVersion.Text = LocalizationManager.GetString("Main_Lbl_TrialRemaining", remainingDays.Value);
                }
                else
                {
                    lblVersion.Text = LocalizationManager.GetString("Main_Status_Trial");
                }
            }
        }
        else
        {
            lblVersion.Text = LocalizationManager.GetString("Main_Status_Trial");
        }
    }

    /// <summary>
    /// 发送待发送的遥测数据
    /// </summary>
    private async Task SendPendingTelemetryAsync()
    {
        try
        {
            TelemetryService.Instance.TrackEvent(TelemetryService.TelemetryEventType.AppStart);
            await TelemetryService.Instance.SendPendingDataAsync();
        }
        catch (Exception ex)
        {
            Log.Warning("[Telemetry] Startup telemetry send failed: {Error}", ex.Message);
        }
    }

    #region EML to PST

    private void BrowseEmlInput_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            txtEmlInputDir.Text = dialog.SelectedPath;
        }
    }

    private void BrowsePstOutput_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            txtPstOutputDir.Text = dialog.SelectedPath;
        }
    }

    private void ScanEmlFolder_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var inputDir = txtEmlInputDir.Text;
            if (!Directory.Exists(inputDir))
            {
                lblEmlEstimatedCount.Text = LocalizationManager.GetString("Main_Msg_DirNotExist");
                return;
            }

            var scanner = new EmLScannerService();
            var files = scanner.ScanEmLFiles(inputDir, true, ".eml").ToList();
            int total = files.Count;

            // 跳过年份、发件人等已归类的文件夹
            int skipCount = 0;
            var dirsToSkip = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (var file in files)
            {
                var dir = Path.GetDirectoryName(file) ?? "";
                var relDir = Path.GetRelativePath(inputDir, dir);
                if (IsClassificationFolder(Path.GetFileName(relDir)))
                {
                    dirsToSkip.Add(dir);
                    skipCount++;
                }
            }

            int actualCount = total - skipCount;
            lblEmlEstimatedCount.Text = LocalizationManager.GetString("Main_Lbl_ScanResult", total, actualCount, skipCount);
        }
        catch (Exception ex)
        {
            lblEmlEstimatedCount.Text = LocalizationManager.GetString("Main_Lbl_ScanFailed", ex.Message);
        }
    }

    private async void ConvertEmlToPst_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            TelemetryService.Instance.TrackEvent(TelemetryService.TelemetryEventType.EmlToPstConversion);

            var inputDir = txtEmlInputDir.Text;
            var outputDir = txtPstOutputDir.Text;

            Log.Information("[EML2PST] ========== 开始 EML 转 PST ==========");
            Log.Information("[EML2PST] 输入目录: {InputDir}", inputDir);
            Log.Information("[EML2PST] 输出目录: {OutputDir}", outputDir);

            lblStatus.Text = LocalizationManager.GetString("Main_Lbl_ScanningEml");

            if (!Directory.Exists(inputDir))
            {
                Log.Error("[EML2PST] 输入目录不存在: {InputDir}", inputDir);
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_EmlInputDirNotExist"),
                    LocalizationManager.GetString("Common_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
                lblStatus.Text = LocalizationManager.GetString("Main_Msg_EmlInputDirNotExist").TrimEnd('！', '!');
                return;
            }

            lblEmlProgress.Text = LocalizationManager.GetString("Main_Lbl_ScanningEml");
            progressEml.Value = 0;

            var scanner = new EmLScannerService();
            var settings = ConfigurationLoader.Load();

            var files = scanner.ScanEmLFiles(inputDir, true, ".eml").ToList();
            int total = files.Count;

            Log.Information("[EML2PST] 扫描完成，找到 {Total} 个 EML 文件", total);

            if (total == 0)
            {
                Log.Warning("[EML2PST] 未找到 EML 文件");
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_NoEmlFiles") + "!",
                    LocalizationManager.GetString("Common_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Information);
                lblStatus.Text = LocalizationManager.GetString("Main_Msg_NoEmlFiles");
                return;
            }

            // 禁用按钮防止重复点击
            var btn = sender as System.Windows.Controls.Button;
            if (btn != null) btn.IsEnabled = false;

            lblEmlProgress.Text = LocalizationManager.GetString("Main_Lbl_FoundEmlFiles", total);
            lblStatus.Text = LocalizationManager.GetString("Main_Lbl_Converting", 0, total);

            // 获取分类选项（在进入后台线程前捕获UI值）
            bool byYear = chkEmlByYear.IsChecked == true;
            bool bySender = chkEmlBySender.IsChecked == true;
            bool byDate = chkEmlByDate.IsChecked == true;

            // 在后台线程执行转换
            var result = await Task.Run(() =>
            {
                int processed = 0;
                int successCount = 0;
                int failCount = 0;
                var parser = new EmailParserService();
                var classifier = new ClassificationService();

                using var pstWriter = new PstWriterService();
                pstWriter.BeginSession();

                // Python os.walk 方式遍历目录，保留目录结构
                var allFiles = Directory.GetFiles(inputDir, "*.eml", SearchOption.AllDirectories);
                var dirGroups = allFiles.GroupBy(f => Path.GetDirectoryName(f) ?? inputDir);

                // 如果勾选了分类选项但没有预定义规则，则动态生成规则
                var effectiveRules = settings.ClassificationRules;
                if ((byYear || bySender || byDate) && !settings.ClassificationRules.Any())
                {
                    // 动态规则会在下面每封邮件处理时生成
                }

                int skipCount = 0;
                int lastUiUpdate = 0;
                foreach (var group in dirGroups)
                {
                    // 计算相对路径
                    var relDir = Path.GetRelativePath(inputDir, group.Key);
                    if (relDir == ".") relDir = "";

                    // 跳过年份、发件人域名等已归类的文件夹
                    var folderName = Path.GetFileName(relDir);
                    if (IsClassificationFolder(folderName))
                    {
                        skipCount += group.Count();
                        continue;
                    }

                    foreach (var file in group)
                    {
                        try
                        {
                            var email = parser.ParseEmail(file);

                            // 优先使用动态分类（按年/按发件人/按月份）
                            if (byYear || bySender || byDate)
                            {
                                var dynamicRules = GenerateDynamicRules(email, byYear, bySender, byDate);
                                foreach (var rule in dynamicRules)
                                {
                                    pstWriter.AddEmailToPst(rule, email, outputDir, relDir);
                                }
                            }
                            else
                            {
                                // 使用预定义规则
                                var results = classifier.ClassifyEmails(new[] { email }, settings.ClassificationRules);
                                foreach (var res in results)
                                {
                                    foreach (var rule in res.MatchedRules)
                                    {
                                        pstWriter.AddEmailToPst(rule, res.Email, outputDir, relDir);
                                    }
                                }
                            }

                            Log.Debug("[EML2PST] 成功处理: {File}", file);
                            successCount++;
                        }
                        catch (Exception ex)
                        {
                            Log.Error(ex, "[EML2PST] 处理失败 {File}: {Error}", file, ex.Message);
                            failCount++;
                        }

                        processed++;

                        // 每处理10封邮件更新一次UI，避免频繁刷新导致冻结
                        if (processed - lastUiUpdate >= 10)
                        {
                            lastUiUpdate = processed;
                            // 更新进度（回到UI线程）
                            System.Windows.Application.Current.Dispatcher.Invoke(() =>
                            {
                                progressEml.Value = (double)processed / total * 100;
                                lblEmlProgress.Text = LocalizationManager.GetString("Main_Lbl_Processing", processed, total);
                                lblStatus.Text = LocalizationManager.GetString("Main_Lbl_Converting", processed, total);
                            });
                        }
                    }
                }

                pstWriter.FinalizeAllPst();

                return (processed, successCount, failCount);
            });

            // 恢复按钮
            if (btn != null) btn.IsEnabled = true;

            lblEmlProgress.Text = LocalizationManager.GetString("Main_Lbl_ConvertCompleteLabel", result.processed);
            lblStatus.Text = LocalizationManager.GetString("Main_Lbl_ConvertCompleteStatus", result.successCount, result.failCount);

            Log.Information("[EML2PST] ========== 转换完成 ==========");
            Log.Information("[EML2PST] 总计: {Total}, 成功: {Success}, 失败: {Fail}", total, result.successCount, result.failCount);
            Log.Information("[EML2PST] PST 文件已保存到: {OutputDir}", outputDir);

            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_ConvertDone", result.successCount, result.failCount, outputDir),
                LocalizationManager.GetString("Common_Complete"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            Log.Error(ex, "[EML2PST] EML to PST 转换失败");
            lblStatus.Text = LocalizationManager.GetString("Main_Lbl_ConvertFailed");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_ConvertFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    #endregion

    #region EML Date Split

    private void BrowseSplitEmlInput_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            txtSplitEmlInputDir.Text = dialog.SelectedPath;
        }
    }

    private void BrowseSplitBefore_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            txtSplitBeforeDir.Text = dialog.SelectedPath;
        }
    }

    private void BrowseSplitAfter_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            txtSplitAfterDir.Text = dialog.SelectedPath;
        }
    }

    private async void SplitEmlByDate_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            TelemetryService.Instance.TrackEvent(TelemetryService.TelemetryEventType.EmlDateSplit);

            var inputDir = txtSplitEmlInputDir.Text;
            var beforeDir = txtSplitBeforeDir.Text;
            var afterDir = txtSplitAfterDir.Text;
            var splitDate = dateSplit.SelectedDate;

            if (!Directory.Exists(inputDir))
            {
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_EmlInputDirNotExist"),
                    LocalizationManager.GetString("Common_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (splitDate == null)
            {
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_SelectSplitDate"),
                    LocalizationManager.GetString("Common_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // Get classification settings
            bool enableClassification = chkEnableClassification.IsChecked == true;
            var patternItem = cboClassificationPattern.SelectedItem as System.Windows.Controls.ComboBoxItem;
            string patternTag = patternItem?.Tag?.ToString() ?? "Year";

            // Create output directories
            Directory.CreateDirectory(beforeDir);
            Directory.CreateDirectory(afterDir);

            lblSplitProgress.Text = LocalizationManager.GetString("Main_Lbl_ScanningEml");
            progressSplit.Value = 0;

            var scanner = new EmLScannerService();
            var parser = new EmailParserService();

            var files = scanner.ScanEmLFiles(inputDir, true, ".eml").ToList();
            int total = files.Count;

            if (total == 0)
            {
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_NoEmlFiles") + "!",
                    LocalizationManager.GetString("Common_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            lblSplitProgress.Text = LocalizationManager.GetString("Main_Lbl_FoundEmlFilesSplit", total);
            AppendLog($"[日期分割] 开始分割，共 {total} 个 EML 文件");
            AppendLog($"[日期分割] 分割日期: {splitDate.Value:yyyy-MM-dd}");
            AppendLog($"[日期分割] 日期之前目录: {beforeDir}");
            AppendLog($"[日期分割] 日期之后目录: {afterDir}");

            int beforeCount = 0;
            int afterCount = 0;
            int processed = 0;
            var emailDetails = new System.Collections.ObjectModel.ObservableCollection<Models.EmailDetailItem>();

            dgEmailDetails.ItemsSource = emailDetails;

            foreach (var file in files)
            {
                try
                {
                    var email = parser.ParseEmail(file);

                    // Get relative directory path to determine if it's a Sent folder
                    var fileDir = Path.GetDirectoryName(file) ?? inputDir;
                    var relDir = Path.GetRelativePath(inputDir, fileDir);
                    if (relDir == ".") relDir = "";

                    // Detect if this is a Sent folder (check for "已发送", "Sent", "发件箱" etc.)
                    bool isSentFolder = IsSentFolder(relDir);

                    // Use SentDate for Sent folder emails, ReceivedDate for others
                    var emailDate = isSentFolder ? email.SentDate : email.ReceivedDate;
                    bool isBefore = emailDate < splitDate.Value.Date;

                    // Create detail item to show both dates and folder type
                    var detailItem = new Models.EmailDetailItem
                    {
                        FileName = Path.GetFileName(file),
                        SentDate = email.SentDate,
                        ReceivedDate = email.ReceivedDate,
                        FromAddress = email.FromAddress,
                        Subject = email.Subject,
                        IsBeforeSplitDate = isBefore,
                        FolderType = isSentFolder
                            ? LocalizationManager.GetString("Main_FolderType_Sent")
                            : LocalizationManager.GetString("Main_FolderType_Inbox")
                    };

                    // Determine destination directory
                    string destDir = isBefore ? beforeDir : afterDir;

                    // Apply classification if enabled - preserve original directory structure
                    if (enableClassification)
                    {
                        var classifiedPath = GetClassifiedDirectory(email, isSentFolder, patternTag, relDir);
                        // Preserve original directory structure and add classification subdirectory
                        if (!string.IsNullOrEmpty(relDir) && relDir != ".")
                        {
                            destDir = Path.Combine(destDir, relDir, classifiedPath);
                        }
                        else
                        {
                            destDir = Path.Combine(destDir, classifiedPath);
                        }
                    }

                    // Create destination directory if it doesn't exist
                    Directory.CreateDirectory(destDir);

                    // Copy file to destination
                    var destFile = Path.Combine(destDir, Path.GetFileName(file));
                    File.Copy(file, destFile, true);

                    if (isBefore)
                        beforeCount++;
                    else
                        afterCount++;

                    // Add to details list
                    Dispatcher.Invoke(() => emailDetails.Add(detailItem));
                }
                catch (Exception ex)
                {
                    Log.Warning("处理失败 {File}: {Error}", file, ex.Message);
                    AppendLog($"[日期分割] 处理失败: {Path.GetFileName(file)}, 错误: {ex.Message}");
                }

                processed++;
                progressSplit.Value = (double)processed / total * 100;
                lblSplitProgress.Text = LocalizationManager.GetString("Main_Lbl_SplitProgress", processed, total, beforeCount, afterCount);
            }

            AppendLog($"[日期分割] 分割完成，日期之前: {beforeCount} 封, 日期之后: {afterCount} 封");
            AppendLog($"[日期分割] 日期之前保存到: {beforeDir}");
            AppendLog($"[日期分割] 日期之后保存到: {afterDir}");
            lblSplitProgress.Text = LocalizationManager.GetString("Main_Lbl_SplitComplete", beforeCount, afterCount);
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_SplitComplete", beforeCount, beforeDir, afterCount, afterDir),
                LocalizationManager.GetString("Common_Complete"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            AppendLog($"[日期分割] 分割失败: {ex.Message}");
            Log.Error(ex, "EML 日期分割失败");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_SplitFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    /// <summary>
    /// 判断目录是否为已发送文件夹
    /// </summary>
    private bool IsSentFolder(string relativeDir)
    {
        if (string.IsNullOrEmpty(relativeDir))
            return false;

        var dirName = relativeDir.ToLowerInvariant();

        // Check for common "Sent" folder names in Chinese and English
        return dirName.Contains("已发送") ||
               dirName.Contains("sent") ||
               dirName.Contains("发件") ||
               dirName.Contains("发件箱") ||
               dirName.Contains("已发") ||
               dirName.Contains("sent items") ||
               dirName.Contains("sentmail");
    }

    /// <summary>
    /// 根据分类模式获取目标目录路径
    /// </summary>
    private string GetClassifiedDirectory(EmailMessage email, bool isSentFolder, string pattern, string originalRelDir)
    {
        // Determine the year to use (sent date for Sent folder, received date for others)
        var dateToUse = isSentFolder ? email.SentDate : email.ReceivedDate;
        var year = dateToUse.Year.ToString();

        // Determine sender and recipient
        var sender = SanitizeFolderName(email.FromAddress);
        var recipient = email.ToAddress.Split(';', StringSplitOptions.RemoveEmptyEntries).FirstOrDefault() ?? "";
        recipient = SanitizeFolderName(recipient);

        // If original directory is a Sent folder, use sender as the person (for sent folder classification)
        // For received emails, use sender domain/address
        string person = !string.IsNullOrEmpty(sender) ? sender : "unknown";

        // Build classified path based on pattern
        return pattern switch
        {
            "Year" => year,
            "Sender" => person,
            "Recipient" => string.IsNullOrEmpty(recipient) ? "unknown" : recipient,
            _ => year
        };
    }

    /// <summary>
    /// 清理文件夹名称，移除非法字符
    /// </summary>
    private string SanitizeFolderName(string name)
    {
        if (string.IsNullOrEmpty(name))
            return "unknown";

        // Replace invalid folder characters but keep the full email address
        var invalidChars = Path.GetInvalidFileNameChars();
        foreach (var c in invalidChars)
        {
            name = name.Replace(c, '_');
        }

        // Limit length
        if (name.Length > 100)
            name = name.Substring(0, 100);

        return name;
    }

    #endregion

    #region PST Extract

    private void BrowsePstInput_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog { Filter = "PST 文件|*.pst|所有文件|*.*", Title = LocalizationManager.GetString("Main_Dialog_SelectPstFile") };
        if (dialog.ShowDialog() == true)
        {
            txtPstInput.Text = dialog.FileName;
        }
    }

    private void BrowseEmlOutput_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            txtEmlOutput.Text = dialog.SelectedPath;
        }
    }

    private async void ExtractPstToEml_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            TelemetryService.Instance.TrackEvent(TelemetryService.TelemetryEventType.PstExtraction);

            var pstPath = txtPstInput.Text;
            var outputDir = txtEmlOutput.Text;

            if (!File.Exists(pstPath))
            {
                AppendLog($"[PST提取] PST 文件不存在: {pstPath}");
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_PstNotExist"),
                    LocalizationManager.GetString("Common_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            AppendLog($"[PST提取] 开始提取邮件...");
            AppendLog($"[PST提取] 源文件: {pstPath}");
            AppendLog($"[PST提取] 输出目录: {outputDir}");
            lblPstProgress.Text = LocalizationManager.GetString("Main_Lbl_Extracting");
            progressPst.Value = 0;

            var progress = new Progress<int>(p => {
                progressPst.Value = p;
                lblPstProgress.Text = LocalizationManager.GetString("Main_Lbl_Extracted", p);
                AppendLog($"[PST提取] 已提取: {p} 封邮件");
            });

            await Task.Run(() => _pstExtractService.ExtractToEml(pstPath, outputDir, progress));

            AppendLog($"[PST提取] 提取完成，共 {progressPst.Value} 封邮件");
            progressPst.Value = 100;
            lblPstProgress.Text = LocalizationManager.GetString("Main_Lbl_ExtractComplete");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_PstExtractComplete", outputDir),
                LocalizationManager.GetString("Common_Complete"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            AppendLog($"[PST提取] 提取失败: {ex.Message}");
            Log.Error(ex, "PST 提取失败");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_PstExtractFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void ExtractPstContacts_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            TelemetryService.Instance.TrackEvent(TelemetryService.TelemetryEventType.PstExtractContacts);

            var pstPath = txtPstInput.Text;
            if (!File.Exists(pstPath))
            {
                AppendLog($"[PST提取] PST 文件不存在: {pstPath}");
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_PstNotExist"),
                    LocalizationManager.GetString("Common_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            AppendLog($"[PST提取] 开始提取联系人...");
            AppendLog($"[PST提取] 源文件: {pstPath}");
            lblPstProgress.Text = LocalizationManager.GetString("Main_Lbl_ExtractingContacts");
            progressPst.Value = 0;

            var contacts = await Task.Run(() => _pstExtractService.ExtractContacts(pstPath));

            AppendLog($"[PST提取] 联系人提取完成，共 {contacts.Count} 个联系人");
            lblPstProgress.Text = LocalizationManager.GetString("Main_Lbl_ContactsComplete");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_ContactsExtracted", contacts.Count),
                LocalizationManager.GetString("Common_Complete"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            AppendLog($"[PST提取] 联系人提取失败: {ex.Message}");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_OperationFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void ExtractPstCalendar_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            TelemetryService.Instance.TrackEvent(TelemetryService.TelemetryEventType.PstExtractCalendar);

            var pstPath = txtPstInput.Text;
            if (!File.Exists(pstPath))
            {
                AppendLog($"[PST提取] PST 文件不存在: {pstPath}");
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_PstNotExist"),
                    LocalizationManager.GetString("Common_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            AppendLog($"[PST提取] 开始提取日历...");
            AppendLog($"[PST提取] 源文件: {pstPath}");
            lblPstProgress.Text = LocalizationManager.GetString("Main_Lbl_ExtractingCalendar");
            progressPst.Value = 0;

            var calendars = await Task.Run(() => _pstExtractService.ExtractCalendar(pstPath));

            AppendLog($"[PST提取] 日历提取完成，共 {calendars.Count} 个事件");
            lblPstProgress.Text = LocalizationManager.GetString("Main_Lbl_ContactsComplete");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_CalendarExtracted", calendars.Count),
                LocalizationManager.GetString("Common_Complete"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            AppendLog($"[PST提取] 日历提取失败: {ex.Message}");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_OperationFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void OrganizePstByYear_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            TelemetryService.Instance.TrackEvent(TelemetryService.TelemetryEventType.PstOrganizeByYear);

            var pstPath = txtClassifyPstInput.Text;
            if (!File.Exists(pstPath))
            {
                AppendLog($"[PST归类] PST 文件不存在: {pstPath}");
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_PstNotExist"),
                    LocalizationManager.GetString("Common_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            var patternItem = cboPstClassificationPattern.SelectedItem as System.Windows.Controls.ComboBoxItem;
            var pattern = patternItem?.Tag?.ToString() ?? "Year";

            AppendLog($"[PST归类] 开始 PST 邮件归类");
            AppendLog($"[PST归类] 源文件: {pstPath}");
            AppendLog($"[PST归类] 分类模式: {pattern}");
            lblClassifyPstProgress.Text = LocalizationManager.GetString("Main_Lbl_ScanningPst");
            progressClassifyPst.Value = 0;

            await Task.Run(() =>
            {
                try
                {
                    var outlookType = Type.GetTypeFromProgID("Outlook.Application");
                    if (outlookType == null)
                    {
                        throw new InvalidOperationException("无法获取 Outlook.Application 类型");
                    }

                    var outlookApp = Activator.CreateInstance(outlookType);
                    var outlookNamespace = outlookApp.GetType().InvokeMember("Session",
                        System.Reflection.BindingFlags.GetProperty, null, outlookApp, null);

                    // 添加 PST 文件到当前会话
                    var pstFolder = outlookNamespace.GetType().InvokeMember("Folders",
                        System.Reflection.BindingFlags.GetProperty, null, outlookNamespace, null);

                    object? addedPst = null;
                    try
                    {
                        addedPst = outlookNamespace.GetType().InvokeMember("AddStore",
                            System.Reflection.BindingFlags.InvokeMethod, null, outlookNamespace, new object[] { pstPath });
                    }
                    catch { }

                    System.Threading.Thread.Sleep(500);

                    // 查找 PST 根文件夹
                    object? rootFolder = null;
                    foreach (object folder in (System.Collections.IEnumerable)pstFolder)
                    {
                        try
                        {
                            var store = folder.GetType().InvokeMember("Store",
                                System.Reflection.BindingFlags.GetProperty, null, folder, null);
                            if (store != null)
                            {
                                var filePath = store.GetType().InvokeMember("FilePath",
                                    System.Reflection.BindingFlags.GetProperty, null, store, null) as string;
                                if (!string.IsNullOrEmpty(filePath) && filePath.Equals(pstPath, StringComparison.OrdinalIgnoreCase))
                                {
                                    rootFolder = folder;
                                    break;
                                }
                            }
                        }
                        catch { }
                    }

                    if (rootFolder == null)
                    {
                        throw new InvalidOperationException("无法打开 PST 文件");
                    }

                    int totalEmails = 0;
                    int processedEmails = 0;
                    var foldersToProcess = new List<object>();

                    // 收集所有文件夹
                    CollectFolders(rootFolder, foldersToProcess);
                    totalEmails = foldersToProcess.Sum(f => GetFolderItemCount(f));
                    AppendLog($"[PST归类] 共发现 {totalEmails} 封邮件，开始处理...");

                    // 处理每个文件夹中的邮件
                    foreach (var folder in foldersToProcess)
                    {
                        ProcessPstFolderForOrganization(folder, pattern, ref processedEmails, totalEmails,
                            () => UpdateClassifyPstProgress(processedEmails, totalEmails));
                    }

                    AppendLog($"[PST归类] 归类完成，共处理 {processedEmails} 封邮件");

                    // 移除添加的 PST store
                    if (addedPst != null)
                    {
                        try
                        {
                            outlookNamespace.GetType().InvokeMember("RemoveStore",
                                System.Reflection.BindingFlags.InvokeMethod, null, outlookNamespace, new object[] { rootFolder });
                        }
                        catch { }
                    }

                    Marshal.ReleaseComObject(rootFolder);
                    Marshal.ReleaseComObject(outlookNamespace);
                    Marshal.ReleaseComObject(outlookApp);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[PST Organize] 归类失败");
                    throw;
                }
            });

            lblClassifyPstProgress.Text = LocalizationManager.GetString("Main_Lbl_ClassifyComplete");
            progressClassifyPst.Value = 100;
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_PstClassifyComplete"),
                LocalizationManager.GetString("Common_Complete"),
                MessageBoxButton.OK, MessageBoxImage.Information);
            Log.Information("[PST Organize] PST 归类完成");
        }
        catch (Exception ex)
        {
            AppendLog($"[PST归类] 归类失败: {ex.Message}");
            Log.Error(ex, "[PST Organize] 归类失败");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_ClassifyFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void CollectFolders(object folder, List<object> folders)
    {
        folders.Add(folder);
        try
        {
            var subFolders = folder.GetType().InvokeMember("Folders",
                System.Reflection.BindingFlags.GetProperty, null, folder, null);
            foreach (object subFolder in (System.Collections.IEnumerable)subFolders)
            {
                // 跳过年份、发件人域名等归类产生的文件夹，避免重复归类
                var subFolderName = subFolder.GetType().InvokeMember("Name",
                    System.Reflection.BindingFlags.GetProperty, null, subFolder, null) as string ?? "";
                if (IsClassificationFolder(subFolderName))
                {
                    continue;
                }
                CollectFolders(subFolder, folders);
            }
        }
        catch { }
    }

    private bool IsYearFolder(string folderName)
    {
        // 年份文件夹：4位数字，1900-2100范围内
        if (folderName.Length == 4 && int.TryParse(folderName, out int year))
        {
            return year >= 1900 && year <= 2100;
        }
        return false;
    }

    private List<ClassificationRule> GenerateDynamicRules(EmailMessage email, bool byYear, bool bySender, bool byDate)
    {
        var rules = new List<ClassificationRule>();

        // 使用统一的时间戳生成PST文件名
        string pstFileName = $"{DateTime.Now:yyyyMMdd-HHmm}.pst";

        // 按年份分类：Inbox/2024, Inbox/2023 等
        if (byYear)
        {
            var yearRule = new ClassificationRule
            {
                Name = $"Year_{email.Year}",
                OutputPstFileName = pstFileName,
                UseFolderTree = true,
                FolderPath = $"Inbox/{email.Year}",
                CombineOperator = RuleCombineOperator.And,
                Conditions = new List<RuleCondition>
                {
                    new RuleCondition { Field = RuleFieldType.Year, Operator = RuleOperator.Equals, Value = email.Year.ToString() }
                }
            };
            rules.Add(yearRule);
        }

        // 按发件人域名分类：Inbox/gmail.com, Inbox/company.com 等
        if (bySender && !string.IsNullOrEmpty(email.FromDomain))
        {
            var senderRule = new ClassificationRule
            {
                Name = $"Sender_{email.FromDomain}",
                OutputPstFileName = pstFileName,
                UseFolderTree = true,
                FolderPath = $"Inbox/{email.FromDomain}",
                CombineOperator = RuleCombineOperator.And,
                Conditions = new List<RuleCondition>
                {
                    new RuleCondition { Field = RuleFieldType.SenderDomain, Operator = RuleOperator.Equals, Value = email.FromDomain }
                }
            };
            rules.Add(senderRule);
        }

        // 按月份分类：Inbox/2024/01, Inbox/2024/02 等
        if (byDate)
        {
            var monthFolder = email.ReceivedDate.ToString("yyyy/MM");
            var dateRule = new ClassificationRule
            {
                Name = $"Date_{email.ReceivedDate:yyyyMM}",
                OutputPstFileName = pstFileName,
                UseFolderTree = true,
                FolderPath = $"Inbox/{monthFolder}",
                CombineOperator = RuleCombineOperator.And,
                Conditions = new List<RuleCondition>
                {
                    new RuleCondition { Field = RuleFieldType.ReceivedDate, Operator = RuleOperator.Equals, Value = email.ReceivedDate.ToString("yyyy-MM-dd") }
                }
            };
            rules.Add(dateRule);
        }

        return rules;
    }

    private bool IsClassificationFolder(string folderName)
    {
        // 归类文件夹识别：跳过年份、发件人域名、日期等归类产生的文件夹
        // 年份：4位数字，1900-2100
        if (IsYearFolder(folderName)) return true;

        // 常见邮件域名（发件人域名文件夹）
        string[] commonDomains = { "gmail", "yahoo", "hotmail", "outlook", "163", "126", "qq", "sina", "sohu", "tom" };
        foreach (var domain in commonDomains)
        {
            if (folderName.Equals(domain, StringComparison.OrdinalIgnoreCase) ||
                folderName.Equals(domain + ".com", StringComparison.OrdinalIgnoreCase) ||
                folderName.Equals(domain + ".cn", StringComparison.OrdinalIgnoreCase))
                return true;
        }

        // 检查是否是邮箱地址格式（包含@且有域名后缀，如 user@domain.com）
        // 这种格式的是发件人/收件人分类文件夹
        if (folderName.Contains("@"))
        {
            var parts = folderName.Split('@');
            if (parts.Length == 2 && parts[1].Contains("."))
            {
                // 看起来像邮箱地址，可能是分类文件夹
                // 但要排除 OST 根文件夹（账户邮箱），它们通常包含 onmicrosoft.com, gmail.com 等
                var domain = parts[1].ToLowerInvariant();
                // 常见免费邮箱域名通常是分类结果
                if (domain.Contains("gmail") || domain.Contains("yahoo") || domain.Contains("hotmail") ||
                    domain.Contains("163.") || domain.Contains("126.") || domain.Contains("qq.") ||
                    domain.Contains("outlook") || domain.Contains("sina") || domain.Contains("sohu") || domain.Contains("tom"))
                    return true;
            }
        }

        // 日期格式文件夹：yyyy-MM, yyyy_MM, yyyyMM 等
        if (folderName.Length >= 6 && folderName.Length <= 10 &&
            (folderName.StartsWith("20") || folderName.StartsWith("19")) &&
            (folderName.Contains("-") || folderName.Contains("_") || folderName.Contains("/")))
            return true;

        return false;
    }

    private int GetFolderItemCount(object folder)
    {
        try
        {
            var items = folder.GetType().InvokeMember("Items",
                System.Reflection.BindingFlags.GetProperty, null, folder, null);
            return (int)items.GetType().InvokeMember("Count",
                System.Reflection.BindingFlags.GetProperty, null, items, null);
        }
        catch
        {
            return 0;
        }
    }

    private void ProcessPstFolderForOrganization(object folder, string pattern, ref int processed, int total, Action updateProgress)
    {
        try
        {
            var folderPath = GetFolderPath(folder);
            var isSentFolder = IsSentFolder(folderPath);
            var folderName = folder.GetType().InvokeMember("Name",
                System.Reflection.BindingFlags.GetProperty, null, folder, null) as string ?? "Unknown";

            // 先检查文件夹本身有多少邮件
            var items = folder.GetType().InvokeMember("Items",
                System.Reflection.BindingFlags.GetProperty, null, folder, null);
            var totalItemCount = (int)items.GetType().InvokeMember("Count",
                System.Reflection.BindingFlags.GetProperty, null, items, null);

            // 检查是否存在分类子文件夹
            var subFolders = folder.GetType().InvokeMember("Folders",
                System.Reflection.BindingFlags.GetProperty, null, folder, null);
            bool hasClassifiedSubFolders = false;
            foreach (object sub in (System.Collections.IEnumerable)subFolders)
            {
                var subName = sub.GetType().InvokeMember("Name",
                    System.Reflection.BindingFlags.GetProperty, null, sub, null) as string ?? "";
                if (IsClassificationFolder(subName))
                {
                    hasClassifiedSubFolders = true;
                    Marshal.ReleaseComObject(sub);
                    break;
                }
                Marshal.ReleaseComObject(sub);
            }
            Marshal.ReleaseComObject(subFolders);

            // 如果存在分类子文件夹，且文件夹本身没有邮件，则跳过
            if (hasClassifiedSubFolders && totalItemCount == 0)
            {
                Log.Information("[PST Organize] 文件夹 {Folder} 已归类（子文件夹存在且无待归类邮件），跳过", folderName);
                AppendLog($"[PST归类] 文件夹 '{folderName}' 已归类，跳过");
                Marshal.ReleaseComObject(items);
                return;
            }

            // 如果存在分类子文件夹但文件夹本身还有邮件，仍然继续归类（可能会重复归类某些邮件）
            if (hasClassifiedSubFolders && totalItemCount > 0)
            {
                AppendLog($"[PST归类] 文件夹 '{folderName}' 有 {totalItemCount} 封邮件待归类，继续处理...");
            }

            Log.Information("[PST Organize] 开始处理文件夹: {Folder}", folderName);

            var count = totalItemCount;

            AppendLog($"[PST归类] 处理文件夹: {folderName}，共 {count} 封邮件");

            for (int i = count; i >= 1; i--)
            {
                try
                {
                    var item = items.GetType().InvokeMember("Item",
                        System.Reflection.BindingFlags.InvokeMethod, null, items, new object[] { i });

                    if (item == null) continue;

                    var itemClass = item.GetType().InvokeMember("MessageClass",
                        System.Reflection.BindingFlags.GetProperty, null, item, null) as string;

                    // 只处理邮件
                    if (itemClass == "IPM.Note" || string.IsNullOrEmpty(itemClass))
                    {
                        // 获取邮件信息
                        var subject = GetProperty(item, "Subject") as string ?? LocalizationManager.GetString("Main_Lbl_NoSubject");
                        var sentOn = GetProperty(item, "SentOn");
                        var receivedTime = GetProperty(item, "ReceivedTime");
                        var from = GetProperty(item, "SenderEmailAddress") as string ?? "";
                        var to = GetProperty(item, "ToEmailAddress") as string ?? "";

                        // 确定使用的日期
                        DateTime emailDate;
                        DateTime tempDate;
                        try
                        {
                            if (isSentFolder && sentOn != null)
                            {
                                tempDate = Convert.ToDateTime(sentOn);
                            }
                            else if (receivedTime != null)
                            {
                                tempDate = Convert.ToDateTime(receivedTime);
                            }
                            else
                            {
                                // 如果没有日期，使用当前日期
                                tempDate = DateTime.Now;
                            }

                            // 验证年份是否合理（1900-2100 之间），不合理则使用当前年份
                            if (tempDate.Year < 1900 || tempDate.Year > 2100)
                            {
                                tempDate = new DateTime(DateTime.Now.Year, tempDate.Month, tempDate.Day);
                            }
                            emailDate = tempDate;
                        }
                        catch
                        {
                            // 如果转换失败，使用当前日期
                            emailDate = DateTime.Now;
                        }
                        var yearStr = emailDate.Year.ToString();
                        string person = "";

                        // 根据模式设置 person
                        if (pattern == "Sender" || pattern == "Recipient")
                        {
                            if (isSentFolder)
                            {
                                // 已发送文件夹使用收件人
                                person = SanitizeFolderName(to.Split(';').FirstOrDefault() ?? "unknown");
                            }
                            else
                            {
                                // 其他文件夹使用发件人
                                person = SanitizeFolderName(from);
                            }
                            if (string.IsNullOrEmpty(person))
                                person = "unknown";
                        }

                        // 获取当前文件夹名称
                        var currentFolderName = folder.GetType().InvokeMember("Name",
                            System.Reflection.BindingFlags.GetProperty, null, folder, null) as string ?? "Unknown";

                        // 构建目标文件夹路径
                        string targetFolderPath;
                        switch (pattern)
                        {
                            case "Year":
                                targetFolderPath = yearStr;
                                break;
                            case "Sender":
                            case "Recipient":
                                targetFolderPath = person;
                                break;
                            default:
                                targetFolderPath = yearStr;
                                break;
                        }

                        // 获取或创建目标文件夹（在当前文件夹下创建）
                        var targetFolder = GetOrCreatePstFolder(folder, targetFolderPath);

                        // 移动邮件到目标文件夹
                        if (targetFolder != null)
                        {
                            item.GetType().InvokeMember("Move",
                                System.Reflection.BindingFlags.InvokeMethod, null, item, new object[] { targetFolder });
                            Log.Debug("[PST Organize] 邮件移动: {Subject} -> {TargetFolder}", subject, targetFolderPath);
                        }

                        processed++;
                        if (processed % 10 == 0)
                        {
                            updateProgress();
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log.Warning("[PST Organize] 处理邮件失败: {Error}", ex.Message);
                }
            }
        }
        catch (Exception ex)
        {
            Log.Warning("[PST Organize] 处理文件夹失败: {Error}", ex.Message);
        }
    }

    private object? GetOrCreatePstFolder(object parentFolder, string folderPath)
    {
        try
        {
            var parts = folderPath.Split('\\');
            object currentFolder = parentFolder;

            foreach (var part in parts)
            {
                if (string.IsNullOrEmpty(part)) continue;

                object? subFolder = null;
                try
                {
                    var folders = currentFolder.GetType().InvokeMember("Folders",
                        System.Reflection.BindingFlags.GetProperty, null, currentFolder, null);

                    subFolder = folders.GetType().InvokeMember("Item",
                        System.Reflection.BindingFlags.InvokeMethod, null, folders, new object[] { part });
                }
                catch { }

                if (subFolder == null)
                {
                    // 创建新文件夹
                    var folders = currentFolder.GetType().InvokeMember("Folders",
                        System.Reflection.BindingFlags.GetProperty, null, currentFolder, null);
                    subFolder = folders.GetType().InvokeMember("Add",
                        System.Reflection.BindingFlags.InvokeMethod, null, folders, new object[] { part });
                }

                if (subFolder != null)
                    currentFolder = subFolder;
            }

            return currentFolder;
        }
        catch (Exception ex)
        {
            Log.Warning("[PST Organize] 创建文件夹失败 {Path}: {Error}", folderPath, ex.Message);
            return null;
        }
    }

    private string GetFolderPath(object folder)
    {
        try
        {
            return folder.GetType().InvokeMember("FolderPath",
                System.Reflection.BindingFlags.GetProperty, null, folder, null) as string ?? "";
        }
        catch
        {
            return "";
        }
    }

    private object? GetProperty(object obj, string propertyName)
    {
        try
        {
            return obj.GetType().InvokeMember(propertyName,
                System.Reflection.BindingFlags.GetProperty, null, obj, null);
        }
        catch
        {
            return null;
        }
    }

    private void UpdatePstProgress(int processed, int total)
    {
        System.Windows.Application.Current.Dispatcher.Invoke(() =>
        {
            progressPst.Value = total > 0 ? (double)processed / total * 100 : 0;
            lblPstProgress.Text = LocalizationManager.GetString("Main_Lbl_Classifying", processed, total);
        });
    }

    private void AppendLog(string message)
    {
        try
        {
            System.Windows.Application.Current.Dispatcher.Invoke(() =>
            {
                var timestamp = DateTime.Now.ToString("HH:mm:ss");
                var logLine = $"[{timestamp}] {message}";
                if (txtLogPanel.Text.Length > 50000)
                    txtLogPanel.Text = "";
                txtLogPanel.Text += logLine + Environment.NewLine;
                txtLogPanel.ScrollToEnd();
            });
        }
        catch { }
    }

    private void ClearLogPanel_Click(object sender, RoutedEventArgs e)
    {
        txtLogPanel.Text = "";
    }

    #endregion

    #region 邮件归类

    private void BrowseClassifyPstInput_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new Microsoft.Win32.OpenFileDialog { Filter = "PST 文件|*.pst|所有文件|*.*", Title = LocalizationManager.GetString("Main_Dialog_SelectPstFile") };
        if (dialog.ShowDialog() == true)
        {
            txtClassifyPstInput.Text = dialog.FileName;
        }
    }

    private void BrowseClassifyOstInput_Click(object sender, RoutedEventArgs e)
    {
        // OST 文件已经加载到 Outlook，不需要选择文件
        System.Windows.MessageBox.Show(
            LocalizationManager.GetString("Main_Msg_OstAutoLoaded"),
            LocalizationManager.GetString("Common_Info"),
            MessageBoxButton.OK, MessageBoxImage.Information);
    }

    private void UpdateClassifyPstProgress(int processed, int total)
    {
        System.Windows.Application.Current.Dispatcher.Invoke(() =>
        {
            progressClassifyPst.Value = total > 0 ? (double)processed / total * 100 : 0;
            lblClassifyPstProgress.Text = LocalizationManager.GetString("Main_Lbl_Classifying", processed, total);
        });
    }

    private void UpdateClassifyOstProgress(int processed, int total)
    {
        System.Windows.Application.Current.Dispatcher.Invoke(() =>
        {
            progressClassifyOst.Value = total > 0 ? (double)processed / total * 100 : 0;
            lblClassifyOstProgress.Text = LocalizationManager.GetString("Main_Lbl_Classifying", processed, total);
        });
    }

    private async void OrganizeOstByYear_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            TelemetryService.Instance.TrackEvent(TelemetryService.TelemetryEventType.OstOrganizeByYear);

            if (cboOstStore.SelectedItem is not OstStoreInfo selectedStore)
            {
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_SelectOstAccount"),
                    LocalizationManager.GetString("Common_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var patternItem = cboOstClassificationPattern.SelectedItem as System.Windows.Controls.ComboBoxItem;
            var pattern = patternItem?.Tag?.ToString() ?? "Year";

            lblClassifyOstProgress.Text = LocalizationManager.GetString("Main_Lbl_ScanningOst");
            progressClassifyOst.Value = 0;

            await Task.Run(() =>
            {
                try
                {
                    var outlookType = Type.GetTypeFromProgID("Outlook.Application");
                    if (outlookType == null)
                    {
                        throw new InvalidOperationException("无法获取 Outlook.Application 类型");
                    }

                    var outlookApp = Activator.CreateInstance(outlookType);
                    var session = outlookApp.GetType().InvokeMember("Session",
                        System.Reflection.BindingFlags.GetProperty, null, outlookApp, null);

                    // 从 ns.Folders 中查找对应的 OST 文件夹
                    var folders = session.GetType().InvokeMember("Folders",
                        System.Reflection.BindingFlags.GetProperty, null, session, null);

                    dynamic? targetFolder = null;
                    foreach (dynamic folder in (System.Collections.IEnumerable)folders)
                    {
                        try
                        {
                            if (folder.Store != null)
                            {
                                var filePath = folder.Store.FilePath as string;
                                bool isOstStore = string.IsNullOrEmpty(filePath) ||
                                    (!string.IsNullOrEmpty(filePath) && filePath.EndsWith(".ost", StringComparison.OrdinalIgnoreCase));

                                if (isOstStore)
                                {
                                    // 处理 "(Exchange)" 字符串占位符的情况
                                    bool isMatch;
                                    if (selectedStore.FilePath == "(Exchange)")
                                    {
                                        // 如果选择的是 Exchange 账户，则匹配任何空路径的 OST store
                                        isMatch = string.IsNullOrEmpty(filePath);
                                    }
                                    else
                                    {
                                        isMatch = !string.IsNullOrEmpty(filePath) &&
                                            Path.GetFullPath(selectedStore.FilePath).Equals(Path.GetFullPath(filePath), StringComparison.OrdinalIgnoreCase);
                                    }

                                    if (isMatch)
                                    {
                                        targetFolder = folder;
                                        AppendLog($"[OST归类] 找到匹配的 OST 文件夹: {folder.Name}");
                                    }
                                    else
                                    {
                                        AppendLog($"[OST归类] 文件夹不匹配: {folder.Name}, 路径: {filePath ?? "(空)"}");
                                        Marshal.ReleaseComObject(folder);
                                    }
                                }
                                else
                                {
                                    Marshal.ReleaseComObject(folder);
                                }
                            }
                            else
                            {
                                Marshal.ReleaseComObject(folder);
                            }
                        }
                        catch { }
                    }

                    if (targetFolder == null)
                    {
                        throw new InvalidOperationException("无法找到对应的 OST 文件夹。请确保 Outlook 中已加载该账户。");
                    }

                    var foldersToProcess = new List<object>();
                    CollectFolders(targetFolder, foldersToProcess);

                    int totalEmails = foldersToProcess.Sum(f => GetFolderItemCount(f));
                    AppendLog($"[OST归类] 找到 {foldersToProcess.Count} 个文件夹，共 {totalEmails} 封邮件");
                    int processedEmails = 0;

                    foreach (var folder in foldersToProcess)
                    {
                        ProcessPstFolderForOrganization(folder, pattern, ref processedEmails, totalEmails,
                            () => UpdateClassifyOstProgress(processedEmails, totalEmails));
                    }

                    Marshal.ReleaseComObject(folders);
                    Marshal.ReleaseComObject(session);
                    Marshal.ReleaseComObject(outlookApp);
                }
                catch (Exception ex)
                {
                    Log.Error(ex, "[OST Organize] 归类失败");
                    throw;
                }
            });

            lblClassifyOstProgress.Text = LocalizationManager.GetString("Main_Lbl_ClassifyComplete");
            progressClassifyOst.Value = 100;
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_OstClassifyComplete"),
                LocalizationManager.GetString("Common_Complete"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            Log.Error(ex, "[OST Organize] 归类失败");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_ClassifyFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    #endregion

    #region OST Extract

    private void BrowseOstOutput_Click(object sender, RoutedEventArgs e)
    {
        using var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            txtOstOutput.Text = dialog.SelectedPath;
        }
    }

    private async void ExtractOstToEml_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            TelemetryService.Instance.TrackEvent(TelemetryService.TelemetryEventType.OstExtraction);

            if (cboOstExtractStore.SelectedItem is not OstStoreInfo selectedStore)
            {
                AppendLog($"[OST提取] 请先选择一个 OST 账户");
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_SelectOstAccount"),
                    LocalizationManager.GetString("Common_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var outputDir = txtOstOutput.Text;
            if (string.IsNullOrWhiteSpace(outputDir))
            {
                AppendLog($"[OST提取] 请选择输出目录");
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_SelectOutputDir"),
                    LocalizationManager.GetString("Common_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            var startDate = dtpOstStartDate.SelectedDate;
            var endDate = dtpOstEndDate.SelectedDate;

            AppendLog($"[OST提取] 开始提取 OST 邮件...");
            AppendLog($"[OST提取] 源文件: {selectedStore.FilePath}");
            AppendLog($"[OST提取] 输出目录: {outputDir}");
            AppendLog($"[OST提取] 日期范围: {(startDate.HasValue ? startDate.Value.ToString("yyyy-MM-dd") : "不限")} - {(endDate.HasValue ? endDate.Value.ToString("yyyy-MM-dd") : "不限")}");
            lblOstProgress.Text = LocalizationManager.GetString("Main_Lbl_CountingMails");
            progressOst.Value = 0;

            var progress = new Progress<(int current, int total)>(p => {
                if (p.total > 0)
                {
                    progressOst.Value = (double)p.current / p.total * 100;
                    lblOstProgress.Text = LocalizationManager.GetString("Main_Lbl_OstExtractedTotal", p.current, p.total);
                    AppendLog($"[OST提取] 已提取: {p.current}/{p.total} 封邮件");
                }
                else
                {
                    lblOstProgress.Text = LocalizationManager.GetString("Main_Lbl_OstExtracted", p.current);
                }
            });

            await Task.Run(() => _ostExtractService.ExtractToEml(selectedStore.FilePath, outputDir, progress, startDate, endDate));

            AppendLog($"[OST提取] 提取完成");
            lblOstProgress.Text = LocalizationManager.GetString("Main_Lbl_ExtractComplete");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_OstExtractComplete", outputDir),
                LocalizationManager.GetString("Common_Complete"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            AppendLog($"[OST提取] 提取失败: {ex.Message}");
            Log.Error(ex, "OST 提取失败");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_OstExtractFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void UpdateOstProgress(int processed, int total)
    {
        System.Windows.Application.Current.Dispatcher.Invoke(() =>
        {
            progressOst.Value = total > 0 ? (double)processed / total * 100 : 0;
            lblOstProgress.Text = LocalizationManager.GetString("Main_Lbl_Classifying", processed, total);
        });
    }

    #endregion

    #region IMAP

    private void cboImapConfig_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        FillImapFieldsFromSelection();
    }

    private void FillImapFieldsFromSelection()
    {
        try
        {
            // 选择配置后自动填充到输入框
            if (cboImapConfig.SelectedItem is KeyValuePair<string, Dictionary<string, string>> selected)
            {
                var config = selected.Value;
                txtImapServer.Text = config.ContainsKey("Server") ? config["Server"] : "";
                txtImapPort.Text = config.ContainsKey("Port") ? config["Port"] : "993";
                txtImapEmail.Text = config.ContainsKey("Email") ? config["Email"] : "";
                txtImapPassword.Password = config.ContainsKey("Password") ? config["Password"] : "";
                AppendLog($"[IMAP] 已加载配置: {selected.Key}, 服务器: {txtImapServer.Text}");
            }
        }
        catch (Exception ex)
        {
            AppendLog($"[IMAP] 加载配置失败: {ex.Message}");
        }
    }

    private void ManageImapConfig_Click(object sender, RoutedEventArgs e)
    {
        // 打开首选项窗口的IMAP配置页
        var prefsWindow = new PreferencesWindow { Owner = this };
        prefsWindow.ShowDialog();
        LoadImapConfigList(); // 重新加载配置列表
    }

    private void LoadImapConfigList()
    {
        cboImapConfig.Items.Clear();
        try
        {
            var appSettingsPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "usersettings.json");
            if (System.IO.File.Exists(appSettingsPath))
            {
                var json = System.IO.File.ReadAllText(appSettingsPath);
                var settings = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(json);

                if (settings != null && settings.TryGetValue("ImapAccounts", out var imapAccountsJson) && !string.IsNullOrEmpty(imapAccountsJson))
                {
                    var accounts = System.Text.Json.JsonSerializer.Deserialize<List<ImapAccountConfig>>(imapAccountsJson);
                    if (accounts != null)
                    {
                        foreach (var account in accounts)
                        {
                            var config = new Dictionary<string, string>
                            {
                                { "Server", account.Server },
                                { "Port", account.Port },
                                { "Email", account.Email },
                                { "Password", account.Password },
                                { "UseSsl", account.UseSsl.ToString() }
                            };
                            cboImapConfig.Items.Add(new KeyValuePair<string, Dictionary<string, string>>(string.IsNullOrEmpty(account.Name) ? account.Email : account.Name, config));
                        }
                        if (cboImapConfig.Items.Count > 0)
                        {
                            cboImapConfig.SelectedIndex = 0;
                            FillImapFieldsFromSelection();
                        }
                    }
                }
            }
        }
        catch { }
    }

    private void BrowseImapOutput_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new System.Windows.Forms.FolderBrowserDialog();
        if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
        {
            txtImapOutputPst.Text = dialog.SelectedPath;
        }
    }

    private async void TestImapConnection_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            TelemetryService.Instance.TrackEvent(TelemetryService.TelemetryEventType.ImapConnectionTest);

            var email = txtImapEmail.Text;
            var password = txtImapPassword.Password;

            if (string.IsNullOrEmpty(email) || string.IsNullOrEmpty(password))
            {
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_EnterEmailPassword"),
                    LocalizationManager.GetString("Common_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            lblImapServerStatus.Text = LocalizationManager.GetString("Main_Lbl_TestingConnection");
            lblImapServerStatus.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(37, 99, 235));

            _imapService.SetCallbacks(
                msg => Dispatcher.Invoke(() => lblImapServerStatus.Text = msg),
                null
            );

            var folders = await _imapService.GetFoldersAsync(email, password);

            // 填充文件夹列表
            lstImapFolders.Items.Clear();
            foreach (var folder in folders)
            {
                lstImapFolders.Items.Add(folder);
            }

            lblImapServerStatus.Text = LocalizationManager.GetString("Main_Lbl_ImapConnected", folders.Count);
            lblImapServerStatus.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(16, 185, 129));
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_ImapConnected", folders.Count),
                LocalizationManager.GetString("Main_Dialog_ConnectionTest"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            lblImapServerStatus.Text = LocalizationManager.GetString("Main_Lbl_ImapConnectFailed", ex.Message);
            lblImapServerStatus.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(239, 68, 68));
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_ImapConnectFailedDetail", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void RefreshImapFolders_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var email = txtImapEmail.Text;
            var password = txtImapPassword.Password;

            if (string.IsNullOrEmpty(email) || string.IsNullOrEmpty(password))
            {
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_EnterEmailPassword"),
                    LocalizationManager.GetString("Common_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            lblImapServerStatus.Text = LocalizationManager.GetString("Main_Lbl_RefreshingFolders");
            _imapService.SetCallbacks(
                msg => Dispatcher.Invoke(() => lblImapServerStatus.Text = msg),
                null
            );

            var folders = await _imapService.GetFoldersAsync(email, password);

            lstImapFolders.Items.Clear();
            foreach (var folder in folders)
            {
                lstImapFolders.Items.Add(folder);
            }

            lblImapServerStatus.Text = LocalizationManager.GetString("Main_Lbl_FoldersFound", folders.Count);
        }
        catch (Exception ex)
        {
            lblImapServerStatus.Text = LocalizationManager.GetString("Main_Lbl_ImapRefreshFailed", ex.Message);
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_ImapRefreshFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void FetchImapToPst_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            TelemetryService.Instance.TrackEvent(TelemetryService.TelemetryEventType.ImapFetchToPst);

            var email = txtImapEmail.Text;
            var password = txtImapPassword.Password;
            var outputPst = txtImapOutputPst.Text;
            var maxEmails = int.TryParse(txtImapMaxEmails.Text, out var m) ? m : 1000;

            if (string.IsNullOrEmpty(email) || string.IsNullOrEmpty(password))
            {
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_EnterEmailPassword"),
                    LocalizationManager.GetString("Common_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 获取选中的文件夹
            var selectedFolders = lstImapFolders.SelectedItems.Cast<string>().ToList();

            lblImapProgress.Text = LocalizationManager.GetString("Main_Lbl_Fetching");
            lblImapServerStatus.Text = "";
            progressImap.Value = 0;

            _imapService.SetCallbacks(
                msg => Dispatcher.Invoke(() => {
                    lblImapProgress.Text = msg;
                    AppendLog(msg);
                }),
                (current, total) => Dispatcher.Invoke(() => {
                    progressImap.Value = (double)current / total * 100;
                    lblImapProgress.Text = LocalizationManager.GetString("Main_Lbl_ImapProgress", current, total);
                })
            );

            AppendLog($"[IMAP->PST] 开始收取邮件到: {outputPst}");
            await _imapService.FetchToPstAsync(email, password, outputPst, maxEmails, null, selectedFolders, default, default,
                chkImapByYear.IsChecked == true,
                chkImapBySender.IsChecked == true,
                chkImapByDate.IsChecked == true);
            AppendLog($"[IMAP->PST] 收取完成");

            lblImapProgress.Text = LocalizationManager.GetString("Main_Lbl_FetchComplete");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_ImapFetchCompletePst", outputPst),
                LocalizationManager.GetString("Common_Complete"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            Log.Error(ex, "IMAP 收取失败");
            AppendLog($"[错误] IMAP 收取失败: {ex.Message}");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_ImapFetchFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void FetchImapToEml_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            TelemetryService.Instance.TrackEvent(TelemetryService.TelemetryEventType.ImapFetchToEml);
            var email = txtImapEmail.Text;
            var password = txtImapPassword.Password;
            var outputDir = txtImapOutputPst.Text;
            var maxEmails = int.TryParse(txtImapMaxEmails.Text, out var m) ? m : 1000;

            if (string.IsNullOrEmpty(email) || string.IsNullOrEmpty(password))
            {
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_EnterEmailPassword"),
                    LocalizationManager.GetString("Common_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 获取选中的文件夹
            var selectedFolders = lstImapFolders.SelectedItems.Cast<string>().ToList();

            lblImapProgress.Text = LocalizationManager.GetString("Main_Lbl_Fetching");
            lblImapServerStatus.Text = "";
            progressImap.Value = 0;

            _imapService.SetCallbacks(
                msg => Dispatcher.Invoke(() => {
                    lblImapProgress.Text = msg;
                    AppendLog($"[IMAP->EML] {msg}");
                }),
                (current, total) => Dispatcher.Invoke(() => {
                    progressImap.Value = (double)current / total * 100;
                    lblImapProgress.Text = LocalizationManager.GetString("Main_Lbl_ImapProgress", current, total);
                })
            );

            AppendLog($"[IMAP->EML] 开始收取邮件到: {outputDir}");
            await _imapService.FetchToEmlAsync(email, password, outputDir, maxEmails, null, selectedFolders, null, default,
                chkImapByYear.IsChecked == true,
                chkImapBySender.IsChecked == true,
                chkImapByDate.IsChecked == true);
            AppendLog($"[IMAP->EML] 收取完成");

            lblImapProgress.Text = LocalizationManager.GetString("Main_Lbl_FetchComplete");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_ImapFetchCompleteEml", outputDir),
                LocalizationManager.GetString("Common_Complete"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            AppendLog($"[IMAP->EML] 收取失败: {ex.Message}");
            Log.Error(ex, "IMAP 收取失败");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_ImapFetchFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    #endregion

    #region Office 365

    private void cboO365Config_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        // 选择配置后自动填充到输入框
        if (cboO365Config.SelectedItem is KeyValuePair<string, Dictionary<string, string>> selected)
        {
            var config = selected.Value;
            if (config.TryGetValue("TenantId", out var tenantId)) txtTenantId.Text = tenantId;
            if (config.TryGetValue("ClientId", out var clientId)) txtClientId.Text = clientId;
            if (config.TryGetValue("Username", out var username)) txtO365Username.Text = username;
            AppendLog($"[O365] 已加载配置: {selected.Key}");
        }
    }

    private void ManageO365Config_Click(object sender, RoutedEventArgs e)
    {
        // 打开首选项窗口的O365配置页
        var prefsWindow = new PreferencesWindow { Owner = this };
        prefsWindow.ShowDialog();
        LoadO365ConfigList(); // 重新加载配置列表
        LoadOstStoreList(); // 加载 OST 账户列表
    }

    /// <summary>
    /// 加载 Outlook 中的 OST 账户列表
    /// </summary>
    private void LoadOstStoreList()
    {
        // 同时加载归类页面和提取页面的 OST 下拉框
        LoadOstStoreComboBox(cboOstStore);
        LoadOstStoreComboBox(cboOstExtractStore);
    }

    private void LoadOstStoreComboBox(System.Windows.Controls.ComboBox comboBox)
    {
        comboBox.Items.Clear();
        try
        {
            AppendLog("[OST] 正在加载 OST 账户...");
            var outlookType = Type.GetTypeFromProgID("Outlook.Application");
            if (outlookType == null)
            {
                AppendLog("[OST] 无法获取 Outlook.Application 类型");
                Log.Warning("[OST] 无法获取 Outlook.Application 类型");
                return;
            }

            // 使用异步方式创建 Outlook 实例，避免阻塞
            var tcs = new System.Threading.Tasks.TaskCompletionSource<object>();
            System.Threading.Timer? timer = null;
            timer = new System.Threading.Timer(_ =>
            {
                tcs.TrySetCanceled();
                timer?.Dispose();
            }, null, 5000, System.Threading.Timeout.Infinite); // 5秒超时

            var outlookAppTask = System.Threading.Tasks.Task.Run(() =>
            {
                try
                {
                    return Activator.CreateInstance(outlookType);
                }
                catch (Exception ex)
                {
                    tcs.TrySetException(ex);
                    throw;
                }
            });

            var timeoutTask = System.Threading.Tasks.Task.WhenAny(outlookAppTask, tcs.Task);
            if (timeoutTask.Result == tcs.Task)
            {
                AppendLog("[OST] 加载 OST 账户超时");
                Log.Warning("[OST] 加载 OST 账户超时");
                timer?.Dispose();
                return;
            }

            timer?.Dispose();
            var outlookApp = outlookAppTask.Result;
            var session = outlookApp.GetType().InvokeMember("Session",
                System.Reflection.BindingFlags.GetProperty, null, outlookApp, null);

            // 使用 Accounts 属性获取所有账户
            var accounts = session.GetType().InvokeMember("Accounts",
                System.Reflection.BindingFlags.GetProperty, null, session, null);

            int accountCount = 0;
            int ostCount = 0;
            foreach (object account in (System.Collections.IEnumerable)accounts)
            {
                accountCount++;
                try
                {
                    var displayName = account.GetType().InvokeMember("DisplayName",
                        System.Reflection.BindingFlags.GetProperty, null, account, null) as string ?? "未知账户";
                    var smtpAddress = account.GetType().InvokeMember("SmtpAddress",
                        System.Reflection.BindingFlags.GetProperty, null, account, null) as string ?? "";

                    AppendLog($"[OST] 账户 {accountCount}: {displayName}, SMTP: {smtpAddress}");
                    Log.Information("[OST] 账户 {Index}: {Name}, SMTP: {SMTP}", accountCount, displayName, smtpAddress);

                    // 获取账户的 DeliveryStore
                    var deliveryStore = account.GetType().InvokeMember("DeliveryStore",
                        System.Reflection.BindingFlags.GetProperty, null, account, null);

                    if (deliveryStore != null)
                    {
                        var filePath = deliveryStore.GetType().InvokeMember("FilePath",
                            System.Reflection.BindingFlags.GetProperty, null, deliveryStore, null) as string;

                        // 判断是否为 OST 存储：路径为空（Exchange 缓存模式）或以 .ost 结尾
                        bool isOstStore = string.IsNullOrEmpty(filePath) ||
                            (!string.IsNullOrEmpty(filePath) && filePath.EndsWith(".ost", StringComparison.OrdinalIgnoreCase));

                        AppendLog($"[OST]   存储路径: {filePath ?? "(空)"}, 是否OST: {isOstStore}");
                        Log.Information("[OST]   存储路径: {Path}, 是OST: {IsOst}", filePath ?? "(空)", isOstStore);

                        if (isOstStore)
                        {
                            ostCount++;
                            comboBox.Items.Add(new OstStoreInfo
                            {
                                DisplayName = string.IsNullOrEmpty(smtpAddress) ? displayName : smtpAddress,
                                FilePath = filePath ?? "(Exchange)"
                            });
                            AppendLog($"[OST]   已添加 OST 账户: {displayName}");
                        }
                        Marshal.ReleaseComObject(deliveryStore);
                    }
                    else
                    {
                        AppendLog($"[OST]   账户 {displayName} 的 DeliveryStore 为 null");
                    }
                    Marshal.ReleaseComObject(account);
                }
                catch (Exception ex)
                {
                    AppendLog($"[OST] 处理账户 {accountCount} 时出错: {ex.Message}");
                    Log.Warning("[OST] 处理账户 {Index} 时出错: {Error}", accountCount, ex.Message);
                }
            }

            AppendLog($"[OST] 遍历完成，共 {accountCount} 个账户，找到 {ostCount} 个 OST 账户");
            Log.Information("[OST] 共遍历了 {Total} 个账户，找到 {Count} 个 OST 账户", accountCount, comboBox.Items.Count);

            Marshal.ReleaseComObject(accounts);
            Marshal.ReleaseComObject(session);
            Marshal.ReleaseComObject(outlookApp);

            if (comboBox.Items.Count > 0)
            {
                comboBox.SelectedIndex = 0;
            }
            else
            {
                AppendLog("[OST] 未找到任何 OST 账户");
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_OstAccountsNotFound"),
                    LocalizationManager.GetString("Common_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            AppendLog($"[OST] 加载 OST 账户列表失败: {ex.Message}");
            Log.Warning("[OST] 加载 OST 账户列表失败: {Error}", ex.Message);
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_OstAccountListFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void LoadO365ConfigList()
    {
        cboO365Config.Items.Clear();
        try
        {
            var appSettingsPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "usersettings.json");
            if (System.IO.File.Exists(appSettingsPath))
            {
                var json = System.IO.File.ReadAllText(appSettingsPath);
                var settings = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(json);

                if (settings != null && settings.TryGetValue("O365Accounts", out var o365AccountsJson) && !string.IsNullOrEmpty(o365AccountsJson))
                {
                    var accounts = System.Text.Json.JsonSerializer.Deserialize<List<O365AccountConfig>>(o365AccountsJson);
                    if (accounts != null)
                    {
                        foreach (var account in accounts)
                        {
                            var config = new Dictionary<string, string>
                            {
                                { "TenantId", account.TenantId },
                                { "ClientId", account.ClientId },
                                { "Username", account.Username }
                            };
                            cboO365Config.Items.Add(new KeyValuePair<string, Dictionary<string, string>>(account.Name, config));
                        }
                        if (cboO365Config.Items.Count > 0)
                            cboO365Config.SelectedIndex = 0;
                    }
                }
            }
        }
        catch { }
    }

    private async void ConnectO365_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            TelemetryService.Instance.TrackEvent(TelemetryService.TelemetryEventType.O365Connection);

            var tenantId = txtTenantId.Text;
            var clientId = txtClientId.Text;

            if (string.IsNullOrEmpty(tenantId) || string.IsNullOrEmpty(clientId))
            {
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_EnterO365Credentials"),
                    LocalizationManager.GetString("Common_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            lblO365Status.Text = LocalizationManager.GetString("Main_Lbl_OpeningBrowser");

            var (connected, errorMessage) = await _o365Service.ConnectInteractiveAsync(tenantId, clientId);

            if (connected)
            {
                lblO365Status.Text = LocalizationManager.GetString("Main_Lbl_O365Connected");
                lblO365Status.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(34, 197, 94));
                AppendLog("[O365] 登录成功");
                UpdateO365DataPanelVisibility(true);
                tabMailManagement.IsEnabled = true;
                RefreshO365Stats_Click(null, null);
            }
            else
            {
                lblO365Status.Text = LocalizationManager.GetString("Main_Lbl_LoginFailed");
                lblO365Status.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(239, 68, 68));
                AppendLog($"[O365] 登录失败: {errorMessage}");
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_LoginFailedDetail", errorMessage),
                    LocalizationManager.GetString("Common_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        catch (Exception ex)
        {
            lblO365Status.Text = LocalizationManager.GetString("Main_Lbl_LoginFailed");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_LoginFailedDetail", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void DisconnectO365_Click(object sender, RoutedEventArgs e)
    {
        _o365Service = new Office365SyncService();
        lblO365Status.Text = LocalizationManager.GetString("Main_Lbl_Disconnected");
        lblO365Status.Foreground = new System.Windows.Media.SolidColorBrush(System.Windows.Media.Color.FromRgb(100, 116, 139));
        AppendLog("[O365] 已断开连接");
        tabMailManagement.IsEnabled = false;

        // 更新 O365 数据标签页的显示状态
        UpdateO365DataPanelVisibility(false);
    }

    private void UpdateO365DataPanelVisibility(bool isLoggedIn)
    {
        if (pnlO365LoginRequired != null)
        {
            pnlO365LoginRequired.Visibility = isLoggedIn ? Visibility.Collapsed : Visibility.Visible;
        }
        if (pnlO365DataContent != null)
        {
            pnlO365DataContent.Visibility = isLoggedIn ? Visibility.Visible : Visibility.Collapsed;
        }
    }

    private async void RefreshO365Stats_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            AppendLog("[O365] 正在获取邮箱统计数据...");
            // 显示加载状态
            lblO365Sent24h.Text = "...";
            lblO365Received24h.Text = "...";
            lblO365NewMail.Text = "...";
            lblO365SentCount.Text = "...";
            lblO365ReceivedCount.Text = "...";
            lblO365ContactCount.Text = "...";

            var stats = await _o365Service.GetMailboxStatsAsync();
            if (stats != null)
            {
                AppendLog($"[O365] 已获取统计数据 - 24h发送: {stats.Sent24h}, 24h接收: {stats.Received24h}, 未读: {stats.UnreadMessages}");
                lblO365Sent24h.Text = stats.Sent24h.ToString();
                lblO365Received24h.Text = stats.Received24h.ToString();
                lblO365NewMail.Text = stats.UnreadMessages.ToString();
                lblO365SentCount.Text = stats.TotalSent.ToString();
                lblO365ReceivedCount.Text = stats.TotalReceived.ToString();
                lblO365ContactCount.Text = stats.ContactCount.ToString();
            }
            else
            {
                AppendLog("[O365] 无法获取统计数据");
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_O365StatsFailed"),
                    LocalizationManager.GetString("Common_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
            }
        }
        catch (Exception ex)
        {
            AppendLog($"[O365] 获取统计数据失败: {ex.Message}");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_O365StatsError", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void ExportO365Contacts_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV 文件 (*.csv)|*.csv",
                FileName = $"contacts_export_{DateTime.Now:yyyyMMdd_HHmmss}.csv",
                Title = LocalizationManager.GetString("Main_Dialog_ExportContacts")
            };

            if (dialog.ShowDialog() == true)
            {
                AppendLog("[O365] 正在导出联系人...");
                var contacts = await _o365Service.GetContactsAsync();

                if (contacts.Count == 0)
                {
                    AppendLog("[O365] 联系人列表为空");
                    System.Windows.MessageBox.Show(
                        LocalizationManager.GetString("Main_Msg_NoContacts"),
                        LocalizationManager.GetString("Main_Dialog_ExportContacts"),
                        MessageBoxButton.OK, MessageBoxImage.Information);
                    return;
                }

                AppendLog($"[O365] 已获取 {contacts.Count} 个联系人，开始写入文件...");
                // 写入 CSV 文件
                var sb = new System.Text.StringBuilder();
                // CSV 表头
                sb.AppendLine("DisplayName,FirstName,LastName,Email,Phone,CompanyName,Department");

                foreach (var contact in contacts)
                {
                    sb.AppendLine($"\"{EscapeCsv(contact.DisplayName)}\",\"{EscapeCsv(contact.FirstName)}\",\"{EscapeCsv(contact.LastName)}\",\"{EscapeCsv(contact.Email)}\",\"{EscapeCsv(contact.Phone)}\",\"{EscapeCsv(contact.CompanyName)}\",\"{EscapeCsv(contact.Department)}\"");
                }

                await System.IO.File.WriteAllTextAsync(dialog.FileName, sb.ToString(), System.Text.Encoding.UTF8);
                AppendLog($"[O365] 导出联系人完成，共 {contacts.Count} 个");

                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_ContactsExported", dialog.FileName, contacts.Count),
                    LocalizationManager.GetString("Main_Dialog_ExportSuccess"),
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_ContactsExportFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void ImportO365Contacts_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "CSV 文件 (*.csv)|*.csv",
                Title = LocalizationManager.GetString("Main_Dialog_ImportContacts")
            };

            if (dialog.ShowDialog() == true)
            {
                AppendLog($"[O365] 正在读取文件: {dialog.FileName}");

                // 显示正在处理提示
                var processingMsg = LocalizationManager.GetString("Main_Lbl_ParsingCsv");
                AppendLog($"[O365] {processingMsg}");

                // 在后台线程解析CSV
                var result = await Task.Run<(List<ContactData> contacts, int skipped, int totalCount, string filePath)>(async () =>
                {
                    var lines = await System.IO.File.ReadAllLinesAsync(dialog.FileName, System.Text.Encoding.UTF8);
                    if (lines.Length <= 1)
                        return (new List<ContactData>(), 0, 0, dialog.FileName);

                    var header = ParseCsvLine(lines[0]);
                    bool isFoxmailFormat = header.Any(h => h.Contains("姓名") && (header.Any(h2 => h2.Contains("昵称")) || header.Any(h2 => h2.Contains("电子邮箱"))));
                    bool isO365Format = header.Any(h => h.Equals("DisplayName", StringComparison.OrdinalIgnoreCase));

                    var contacts = new List<ContactData>();
                    int skipped = 0;

                    for (int i = 1; i < lines.Length; i++)
                    {
                        if (string.IsNullOrWhiteSpace(lines[i])) continue;

                        var fields = ParseCsvLine(lines[i]);
                        ContactData contact;

                        if (isFoxmailFormat)
                        {
                            contact = new ContactData
                            {
                                DisplayName = fields.Count > 0 ? fields[0] : "",
                                Email = fields.Count > 2 ? fields[2] : "",
                                Phone = fields.Count > 3 ? fields[3] : "",
                                CompanyName = fields.Count > 5 ? fields[5] : "",
                                Department = fields.Count > 6 ? fields[6] : ""
                            };
                            var nameParts = contact.DisplayName.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
                            if (nameParts.Length >= 2)
                            {
                                contact.LastName = nameParts[0];
                                contact.FirstName = string.Join("", nameParts.Skip(1));
                            }
                            else if (nameParts.Length == 1)
                            {
                                contact.LastName = nameParts[0];
                            }
                        }
                        else
                        {
                            contact = new ContactData
                            {
                                DisplayName = fields.Count > 0 ? fields[0] : "",
                                FirstName = fields.Count > 1 ? fields[1] : "",
                                LastName = fields.Count > 2 ? fields[2] : "",
                                Email = fields.Count > 3 ? fields[3] : "",
                                Phone = fields.Count > 4 ? fields[4] : "",
                                CompanyName = fields.Count > 5 ? fields[5] : "",
                                Department = fields.Count > 6 ? fields[6] : ""
                            };
                        }

                        if (!string.IsNullOrEmpty(contact.Email) || !string.IsNullOrEmpty(contact.DisplayName))
                        {
                            contacts.Add(contact);
                        }
                        else
                        {
                            skipped++;
                        }
                    }

                    return (contacts, skipped, lines.Length - 1, dialog.FileName);
                });

                var contacts = result.contacts;
                var skipped = result.skipped;

                AppendLog($"[O365] 检测到格式: {(skipped > 0 ? "Foxmail" : "O365")}, 共 {result.totalCount} 条记录");
                AppendLog($"[O365] 解析完成，有效记录: {contacts.Count}, 跳过: {skipped}");

                if (contacts.Count == 0)
                {
                    System.Windows.MessageBox.Show(
                        LocalizationManager.GetString("Main_Msg_NoValidContactData"),
                        LocalizationManager.GetString("Main_Dialog_ImportContacts"),
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                // 显示进度
                var progress = new Progress<int>(p =>
                {
                    AppendLog($"[O365] 正在导入... {p}%");
                });

                int imported = await _o365Service.ImportContactsAsync(contacts, progress);
                AppendLog($"[O365] 导入完成，成功: {imported}");

                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_ContactsImported", dialog.FileName, contacts.Count, imported),
                    LocalizationManager.GetString("Main_Dialog_ImportSuccess"),
                    MessageBoxButton.OK, MessageBoxImage.Information);

                // 刷新统计数据
                RefreshO365Stats_Click(null, null);
            }
        }
        catch (Exception ex)
        {
            AppendLog($"[O365] 导入失败: {ex.Message}");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_ContactsImportFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private static string EscapeCsv(string value)
    {
        if (string.IsNullOrEmpty(value)) return "";
        return value.Replace("\"", "\"\"").Replace("\n", "").Replace("\r", "");
    }

    private static List<string> ParseCsvLine(string line)
    {
        var fields = new List<string>();
        var current = "";
        bool inQuotes = false;

        for (int i = 0; i < line.Length; i++)
        {
            char c = line[i];
            if (c == '"')
            {
                if (inQuotes && i + 1 < line.Length && line[i + 1] == '"')
                {
                    current += '"';
                    i++;
                }
                else
                {
                    inQuotes = !inQuotes;
                }
            }
            else if (c == ',' && !inQuotes)
            {
                fields.Add(current);
                current = "";
            }
            else
            {
                current += c;
            }
        }
        fields.Add(current);
        return fields;
    }

    private void DownloadFoxmailTemplate_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV 文件 (*.csv)|*.csv",
                FileName = "foxmail_contact_template.csv",
                Title = LocalizationManager.GetString("Main_Dialog_DownloadFoxmailTemplate")
            };

            if (dialog.ShowDialog() == true)
            {
                var template = @"姓名,昵称,电子邮箱,手机,电话,公司,部门,职务,地址,邮编,备注
张三,小张,zhangsan@example.com,13800138000,010-12345678,示例公司,技术部,工程师,北京市朝阳区,100000,这是一个备注
李四,小李,lisi@example.com,13900139000,021-87654321,示例公司,销售部,经理,上海市浦东新区,200000,销售经理";

                System.IO.File.WriteAllText(dialog.FileName, template, System.Text.Encoding.UTF8);
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_FoxmailTemplateSaved", dialog.FileName),
                    LocalizationManager.GetString("Main_Dialog_DownloadSuccess"),
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_TemplateDownloadFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void DownloadO365Template_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV 文件 (*.csv)|*.csv",
                FileName = "o365_contact_template.csv",
                Title = LocalizationManager.GetString("Main_Dialog_DownloadO365Template")
            };

            if (dialog.ShowDialog() == true)
            {
                var template = @"DisplayName,FirstName,LastName,Email,Phone,CompanyName,Department
张三,三,张,zhangsan@example.com,13800138000,示例公司,技术部
李四,四,李,lisi@example.com,13900139000,示例公司,销售部";

                System.IO.File.WriteAllText(dialog.FileName, template, System.Text.Encoding.UTF8);
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_O365TemplateSaved", dialog.FileName),
                    LocalizationManager.GetString("Main_Dialog_DownloadSuccess"),
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_TemplateDownloadFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void BrowseContactSource_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "CSV 文件 (*.csv)|*.csv|所有文件 (*.*)|*.*",
                Title = LocalizationManager.GetString("Main_Dialog_SelectContactFile")
            };

            if (dialog.ShowDialog() == true)
            {
                txtContactSourceFile.Text = dialog.FileName;
            }
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_SelectFileFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void ConvertContactFile_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (string.IsNullOrEmpty(txtContactSourceFile.Text) || !System.IO.File.Exists(txtContactSourceFile.Text))
            {
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_SelectSourceFile"),
                    LocalizationManager.GetString("Common_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            // 读取源文件获取字段列表
            var lines = System.IO.File.ReadAllLines(txtContactSourceFile.Text, System.Text.Encoding.UTF8);
            if (lines.Length < 2)
            {
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_CsvEmpty"),
                    LocalizationManager.GetString("Common_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var headerFields = ParseCsvLine(lines[0]);
            var targetFields = new List<string> { "DisplayName", "FirstName", "LastName", "Email", "Phone", "CompanyName", "Department" };

            // 显示字段映射对话框
            var mappingDialog = new ContactMappingDialog();
            mappingDialog.Initialize(headerFields, targetFields);
            mappingDialog.Owner = this;

            if (mappingDialog.ShowDialog() != true)
                return;

            // 获取用户配置的映射
            var fieldMapping = new Dictionary<string, int>();
            for (int i = 0; i < mappingDialog.MappingItems.Count; i++)
            {
                var item = mappingDialog.MappingItems[i];
                var sourceField = item.SelectedSourceField;
                if (!string.IsNullOrEmpty(sourceField))
                {
                    var idx = headerFields.IndexOf(sourceField);
                    if (idx >= 0)
                        fieldMapping[item.TargetField] = idx;
                }
            }

            // 选择输出文件
            var saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "CSV 文件 (*.csv)|*.csv",
                FileName = $"contacts_converted_{DateTime.Now:yyyyMMdd_HHmmss}.csv",
                Title = LocalizationManager.GetString("Main_Dialog_SaveConvertedFile")
            };

            if (saveDialog.ShowDialog() == true)
            {
                var conversionType = cmbContactConversionType.SelectedIndex;
                AppendLog($"[转换] 开始转换，请稍候...");
                await Task.Run(() => ConvertWithMapping(lines, saveDialog.FileName, fieldMapping, conversionType));
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_ContactsConverted", txtContactSourceFile.Text, saveDialog.FileName),
                    LocalizationManager.GetString("Main_Dialog_ConvertSuccess"),
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_ConvertContactsFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void ConvertWithMapping(string[] lines, string outputFile, Dictionary<string, int> fieldMapping, int conversionType)
    {
        AppendLog($"[转换] 开始转换文件...");
        AppendLog($"[转换] 字段映射: {string.Join(", ", fieldMapping.Select(kv => $"{kv.Key}←{lines[0].Split(',')[kv.Value]}"))}");
        AppendLog($"[转换] 输出文件: {outputFile}");
        AppendLog($"[转换] 转换类型: {(conversionType == 0 || conversionType == 2 || conversionType == 4 ? "O365格式" : "中文格式")}");

        var sb = new System.Text.StringBuilder();

        if (conversionType == 0 || conversionType == 2 || conversionType == 4)
        {
            // 输出 O365 格式
            sb.AppendLine("DisplayName,FirstName,LastName,Email,Phone,CompanyName,Department");
            for (int i = 1; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                var fields = ParseCsvLine(lines[i]);
                var contact = MapWithManualMapping(fields, fieldMapping);
                AppendLog($"[转换] 处理第{i}行: 姓名={contact.DisplayName}, 邮箱={contact.Email}");
                sb.AppendLine($"\"{EscapeCsv(contact.DisplayName)}\",\"{EscapeCsv(contact.FirstName)}\",\"{EscapeCsv(contact.LastName)}\",\"{EscapeCsv(contact.Email)}\",\"{EscapeCsv(contact.Phone)}\",\"{EscapeCsv(contact.CompanyName)}\",\"{EscapeCsv(contact.Department)}\"");
            }
        }
        else
        {
            // 输出中文格式
            sb.AppendLine("显示名,姓,名,邮箱,电话,公司,部门");
            for (int i = 1; i < lines.Length; i++)
            {
                if (string.IsNullOrWhiteSpace(lines[i])) continue;
                var fields = ParseCsvLine(lines[i]);
                var contact = MapWithManualMapping(fields, fieldMapping);
                AppendLog($"[转换] 处理第{i}行: 姓名={contact.DisplayName}, 邮箱={contact.Email}");
                sb.AppendLine($"\"{EscapeCsv(contact.DisplayName)}\",\"{EscapeCsv(contact.LastName)}\",\"{EscapeCsv(contact.FirstName)}\",\"{EscapeCsv(contact.Email)}\",\"{EscapeCsv(contact.Phone)}\",\"{EscapeCsv(contact.CompanyName)}\",\"{EscapeCsv(contact.Department)}\"");
            }
        }

        AppendLog($"[转换] 转换完成，共处理 {lines.Length - 1} 条记录");
        System.IO.File.WriteAllText(outputFile, sb.ToString(), System.Text.Encoding.UTF8);
    }

    private ContactData MapWithManualMapping(List<string> fields, Dictionary<string, int> mapping)
    {
        var contact = new ContactData();

        // 获取姓名
        if (mapping.TryGetValue("DisplayName", out int nameIdx) && nameIdx >= 0 && nameIdx < fields.Count)
        {
            contact.DisplayName = fields[nameIdx].Trim();
            var nameParts = contact.DisplayName.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            if (nameParts.Length >= 2)
            {
                contact.LastName = nameParts[0];
                contact.FirstName = string.Join("", nameParts.Skip(1));
            }
            else if (nameParts.Length == 1)
            {
                contact.LastName = nameParts[0];
            }
        }

        // 获取名
        if (mapping.TryGetValue("FirstName", out int firstIdx) && firstIdx >= 0 && firstIdx < fields.Count)
            contact.FirstName = fields[firstIdx].Trim();

        // 获取姓
        if (mapping.TryGetValue("LastName", out int lastIdx) && lastIdx >= 0 && lastIdx < fields.Count)
            contact.LastName = fields[lastIdx].Trim();

        // 获取邮箱
        if (mapping.TryGetValue("Email", out int emailIdx) && emailIdx >= 0 && emailIdx < fields.Count)
            contact.Email = fields[emailIdx].Trim();

        // 获取电话
        if (mapping.TryGetValue("Phone", out int phoneIdx) && phoneIdx >= 0 && phoneIdx < fields.Count)
            contact.Phone = fields[phoneIdx].Trim();

        // 获取公司
        if (mapping.TryGetValue("CompanyName", out int companyIdx) && companyIdx >= 0 && companyIdx < fields.Count)
            contact.CompanyName = fields[companyIdx].Trim();

        // 获取部门
        if (mapping.TryGetValue("Department", out int deptIdx) && deptIdx >= 0 && deptIdx < fields.Count)
            contact.Department = fields[deptIdx].Trim();

        return contact;
    }

    private void BrowseExportPath_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var dialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "PST 文件 (*.pst)|*.pst",
                FileName = $"{DateTime.Now:yyyyMMdd-HHmm}.pst",
                Title = LocalizationManager.GetString("Main_Dialog_SavePstFile")
            };

            if (dialog.ShowDialog() == true)
            {
                txtExportPstPath.Text = dialog.FileName;
                AppendLog($"[导出] 已选择输出路径: {dialog.FileName}");
            }
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_SelectPathFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void chkExportAllMail_Click(object sender, RoutedEventArgs e)
    {
        var isAllMail = chkExportAllMail.IsChecked == true;
        dpExportStartDate.IsEnabled = !isAllMail;
        dpExportEndDate.IsEnabled = !isAllMail;
    }

    private async void ExportO365MailByDateRange_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            if (string.IsNullOrEmpty(txtExportPstPath.Text))
            {
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_SelectExportPath"),
                    LocalizationManager.GetString("Common_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            var startDate = DateTime.MinValue;
            var endDate = DateTime.MaxValue;

            // 如果没有勾选全部邮件，则使用日期选择器的值
            if (chkExportAllMail.IsChecked != true)
            {
                if (dpExportStartDate.SelectedDate.HasValue)
                    startDate = dpExportStartDate.SelectedDate.Value;
                if (dpExportEndDate.SelectedDate.HasValue)
                    endDate = dpExportEndDate.SelectedDate.Value.AddDays(1).AddSeconds(-1);
            }

            var rangeText = chkExportAllMail.IsChecked == true ? LocalizationManager.GetString("Main_Lbl_AllMails") : $"{startDate:yyyy-MM-dd} 至 {endDate:yyyy-MM-dd}";
            AppendLog($"[导出] 开始导出邮件，时间范围: {rangeText}");
            lblExportStatus.Text = LocalizationManager.GetString("Main_Lbl_ExportingMails");
            progressExport.Value = 0;

            // 提取 UI 值到局部变量，避免跨线程访问
            var exportAllMail = chkExportAllMail.IsChecked == true;
            var outputPath = txtExportPstPath.Text;

            await Task.Run(async () =>
            {
                try
                {
                    // 获取所有邮件文件夹
                    AppendLog($"[导出] 正在获取邮件文件夹列表...");
                    var folders = await _o365Service.GetAllMailFoldersAsync();
                    AppendLog($"[导出] 共找到 {folders.Count} 个文件夹");

                    // 获取邮箱中的邮件（从所有文件夹）
                    var messages = new List<(Microsoft.Graph.Models.Message msg, string folderName, string folderId)>();
                    var pageSize = 50;

                    foreach (var folder in folders)
                    {
                        AppendLog($"[导出] 正在获取文件夹 '{folder.DisplayName}' 中的邮件...");
                        var messagePage = await _o365Service.GetMessagesFromFolderAsync(folder.Id, startDate, endDate, pageSize);

                        while (messagePage?.Value != null)
                        {
                            foreach (var msg in messagePage.Value)
                            {
                                messages.Add((msg, folder.DisplayName, folder.Id));
                            }

                            if (messagePage.OdataNextLink != null)
                            {
                                messagePage = await _o365Service.GetMessagesFromFolderNextPageAsync(messagePage.OdataNextLink);
                            }
                            else
                            {
                                break;
                            }
                        }
                    }

                    AppendLog($"[导出] 共找到 {messages.Count} 封邮件，开始写入 PST 文件...");

                    // 使用 PstWriterService 写入 PST
                    var pstWriter = new PstWriterService();
                    pstWriter.BeginSession();

                    // 读取 UI 设置
                    var exportByYearMonth = System.Windows.Application.Current.Dispatcher.Invoke(() =>
                        chkExportByYearMonth.IsChecked == true);

                    int exported = 0;
                    int failed = 0;

                    for (int i = 0; i < messages.Count; i++)
                    {
                        var (msg, folderName, folderId) = messages[i];
                        try
                        {
                            var emailMsg = ConvertGraphMessageToEmailMessage(msg);
                            if (emailMsg != null && emailMsg.RawMimeMessage != null)
                            {
                                // 根据文件夹名称创建规则，保留原始文件夹结构
                                // 按年月归类时: mail1/2026/04, 不按年月时: mail1
                                var folderPath = exportByYearMonth
                                    ? $"{folderName}/{{year}}/{{month}}"
                                    : folderName;

                                var rule = new ClassificationRule
                                {
                                    Name = "O365导出",
                                    OutputPstFileName = System.IO.Path.GetFileName(outputPath),
                                    UseFolderTree = true,
                                    FolderPath = folderPath
                                };

                                pstWriter.AddEmailToPst(rule, emailMsg, System.IO.Path.GetDirectoryName(outputPath) ?? "", "");
                                exported++;
                            }
                            else
                            {
                                failed++;
                            }
                        }
                        catch (Exception ex)
                        {
                            failed++;
                            Log.Warning("[导出] 邮件导出失败: {Subject}, {Error}", msg.Subject, ex.Message);
                        }

                        // 更新进度
                        var progress = (int)((i + 1) * 100.0 / messages.Count);
                        System.Windows.Application.Current.Dispatcher.Invoke(() =>
                        {
                            progressExport.Value = progress;
                            lblExportStatus.Text = LocalizationManager.GetString("Main_Lbl_O365Exporting", i + 1, messages.Count);
                        });
                    }

                    pstWriter.Dispose();

                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        lblExportStatus.Text = LocalizationManager.GetString("Main_Lbl_O365ExportComplete", exported, failed);
                        progressExport.Value = 100;
                        AppendLog($"[导出] 完成！成功: {exported}, 失败: {failed}");
                    });
                }
                catch (Exception ex)
                {
                    System.Windows.Application.Current.Dispatcher.Invoke(() =>
                    {
                        lblExportStatus.Text = LocalizationManager.GetString("Main_Lbl_O365ExportFailed");
                        AppendLog($"[导出] 导出失败: {ex.Message}");
                    });
                    Log.Error(ex, "[导出] 导出过程出错");
                }
            });
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_O365ExportFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private EmailMessage? ConvertGraphMessageToEmailMessage(Microsoft.Graph.Models.Message msg)
    {
        try
        {
            // 先提取所有字符串数据，避免后续访问 Graph 对象
            var subject = msg.Subject ?? "";
            var toAddress = string.Join("; ", msg.ToRecipients?.Select(r => r.EmailAddress?.Address) ?? Array.Empty<string>());
            var ccAddress = string.Join("; ", msg.CcRecipients?.Select(r => r.EmailAddress?.Address) ?? Array.Empty<string>());
            var fromAddress = msg.From?.EmailAddress?.Address ?? "";
            var fromName = msg.From?.EmailAddress?.Name ?? "";
            var receivedDate = msg.ReceivedDateTime?.DateTime ?? DateTime.Now;
            var sentDate = msg.SentDateTime?.DateTime ?? DateTime.Now;
            var sentDateString = msg.SentDateTime?.DateTime.ToString("yyyy-MM-dd HH:mm:ss") ?? "";
            var bodyContent = msg.Body?.Content ?? "";
            var parentFolderId = msg.ParentFolderId ?? "";

            var emailMsg = new EmailMessage
            {
                Subject = subject,
                ToAddress = toAddress,
                CcAddress = ccAddress,
                FromAddress = fromAddress,
                FromName = fromName,
                ParentFolderId = parentFolderId,
                ReceivedDate = receivedDate,
                SentDate = sentDate,
                SentDateString = sentDateString,
                BodyHtml = bodyContent,
                BodyText = bodyContent
            };

            // 构建 MimeMessage
            if (!string.IsNullOrEmpty(bodyContent))
            {
                var mimeMessage = new MimeMessage();
                mimeMessage.Subject = subject;

                if (!string.IsNullOrEmpty(fromAddress))
                {
                    // 使用 "Display Name <email@domain.com>" 格式
                    var fromFormatted = !string.IsNullOrEmpty(fromName)
                        ? $"{fromName} <{fromAddress}>"
                        : fromAddress;
                    mimeMessage.From.Add(MailboxAddress.Parse(fromFormatted));
                }

                if (!string.IsNullOrEmpty(toAddress))
                {
                    foreach (var addr in toAddress.Split(';', StringSplitOptions.RemoveEmptyEntries))
                    {
                        mimeMessage.To.Add(MailboxAddress.Parse(addr.Trim()));
                    }
                }

                if (!string.IsNullOrEmpty(ccAddress))
                {
                    foreach (var addr in ccAddress.Split(';', StringSplitOptions.RemoveEmptyEntries))
                    {
                        mimeMessage.Cc.Add(MailboxAddress.Parse(addr.Trim()));
                    }
                }

                mimeMessage.Date = new DateTimeOffset(sentDate);
                mimeMessage.Body = new TextPart("html")
                {
                    Text = bodyContent
                };

                emailMsg.RawMimeMessage = mimeMessage;
            }

            return emailMsg;
        }
        catch (Exception ex)
        {
            Log.Warning("[导出] 转换邮件失败: {Error}", ex.Message);
            return null;
        }
    }

    private void ImportO365Emails_Click(object sender, RoutedEventArgs e)
    {
        // 已移除导入功能
    }

    private (string? tenantId, string? clientId) LoadO365Settings()
    {
        try
        {
            var appSettingsPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "usersettings.json");
            if (System.IO.File.Exists(appSettingsPath))
            {
                var json = System.IO.File.ReadAllText(appSettingsPath);
                var settings = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(json);
                if (settings != null)
                {
                    settings.TryGetValue("DefaultTenantId", out var tenantId);
                    settings.TryGetValue("DefaultClientId", out var clientId);
                    return (tenantId, clientId);
                }
            }
        }
        catch { }
        return (null, null);
    }

    private void SelectPstFile_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            AppendLog("[O365] 打开 PST 文件选择对话框");
            Log.Information("[O365] 打开 PST 文件选择对话框");
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "PST 文件 (*.pst)|*.pst",
                Title = LocalizationManager.GetString("Main_Dialog_SelectPstFile")
            };

            if (dialog.ShowDialog() == true)
            {
                _selectedPstFilePath = dialog.FileName;
                lblPstFilePath.Text = _selectedPstFilePath;
                AppendLog($"[O365] 已选择 PST 文件: {_selectedPstFilePath}");
                Log.Information($"[O365] 已选择 PST 文件: {_selectedPstFilePath}");
            }
            else
            {
                AppendLog("[O365] 用户取消选择 PST 文件");
                Log.Information("[O365] 用户取消选择 PST 文件");
            }
        }
        catch (Exception ex)
        {
            AppendLog($"[O365] 选择 PST 文件失败: {ex.Message}");
            Log.Error(ex, "[O365] 选择 PST 文件失败");
        }
    }

    private void SelectEmlFolder_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            AppendLog("[O365] 打开 EML 目录选择对话框");
            Log.Information("[O365] 打开 EML 目录选择对话框");
            using var dialog = new System.Windows.Forms.FolderBrowserDialog();
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                _selectedEmlFolderPath = dialog.SelectedPath;
                lblEmlFolderPath.Text = _selectedEmlFolderPath;
                AppendLog($"[O365] 已选择 EML 目录: {_selectedEmlFolderPath}");
                Log.Information($"[O365] 已选择 EML 目录: {_selectedEmlFolderPath}");
            }
            else
            {
                AppendLog("[O365] 用户取消选择 EML 目录");
                Log.Information("[O365] 用户取消选择 EML 目录");
            }
        }
        catch (Exception ex)
        {
            AppendLog($"[O365] 选择 EML 目录失败: {ex.Message}");
            Log.Error(ex, "[O365] 选择 EML 目录失败");
        }
    }

    private async void SyncPstToO365_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            AppendLog("[O365] 开始同步 PST 到 O365");
            Log.Information("[O365] 开始同步 PST 到 O365");
            TelemetryService.Instance.TrackEvent(TelemetryService.TelemetryEventType.O365SyncPst);

            if (string.IsNullOrEmpty(_selectedPstFilePath))
            {
                AppendLog("[O365] 未选择 PST 文件");
                Log.Warning("[O365] 未选择 PST 文件");
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_SelectPstFile"),
                    LocalizationManager.GetString("Common_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!System.IO.File.Exists(_selectedPstFilePath))
            {
                AppendLog($"[O365] PST 文件不存在: {_selectedPstFilePath}");
                Log.Warning($"[O365] PST 文件不存在: {_selectedPstFilePath}");
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_PstNotExist"),
                    LocalizationManager.GetString("Common_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            AppendLog($"[O365] 开始同步 PST 文件: {_selectedPstFilePath}");
            Log.Information($"[O365] 开始同步 PST 文件: {_selectedPstFilePath}");
            lblO365Status.Text = LocalizationManager.GetString("Main_Lbl_SyncingPst");
            var progress = new Progress<int>(p => {
                lblO365Status.Text = LocalizationManager.GetString("Main_Lbl_SyncProgress", p);
                AppendLog($"[O365] PST 同步进度: {p}%");
            });

            await _o365Service.SyncPstToO365Async(_selectedPstFilePath, "Imported", progress);

            lblO365Status.Text = LocalizationManager.GetString("Main_Lbl_SyncComplete");
            AppendLog("[O365] PST 同步完成");
            Log.Information("[O365] PST 同步完成");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Lbl_SyncComplete"),
                LocalizationManager.GetString("Common_Complete"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            AppendLog($"[O365] PST 同步失败: {ex.Message}");
            Log.Error(ex, "[O365] PST 同步失败");
            lblO365Status.Text = LocalizationManager.GetString("Main_Lbl_O365ExportFailed");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_SyncFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private async void SyncEmlToO365_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            AppendLog("[O365] 开始同步 EML 到 O365");
            Log.Information("[O365] 开始同步 EML 到 O365");
            TelemetryService.Instance.TrackEvent(TelemetryService.TelemetryEventType.O365SyncEml);

            if (string.IsNullOrEmpty(_selectedEmlFolderPath))
            {
                AppendLog("[O365] 未选择 EML 目录");
                Log.Warning("[O365] 未选择 EML 目录");
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_SelectEmlFolder"),
                    LocalizationManager.GetString("Common_Info"),
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (!System.IO.Directory.Exists(_selectedEmlFolderPath))
            {
                AppendLog($"[O365] EML 目录不存在: {_selectedEmlFolderPath}");
                Log.Warning($"[O365] EML 目录不存在: {_selectedEmlFolderPath}");
                System.Windows.MessageBox.Show(
                    LocalizationManager.GetString("Main_Msg_EmlFolderNotExist"),
                    LocalizationManager.GetString("Common_Error"),
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            AppendLog($"[O365] 开始同步 EML 目录: {_selectedEmlFolderPath}");
            Log.Information($"[O365] 开始同步 EML 目录: {_selectedEmlFolderPath}");
            lblO365Status.Text = LocalizationManager.GetString("Main_Lbl_SyncingEml");
            var progress = new Progress<int>(p => {
                lblO365Status.Text = LocalizationManager.GetString("Main_Lbl_SyncProgress", p);
                AppendLog($"[O365] EML 同步进度: {p}%");
            });

            await _o365Service.SyncEmlFolderToO365Async(_selectedEmlFolderPath, "Imported", progress);

            lblO365Status.Text = LocalizationManager.GetString("Main_Lbl_SyncComplete");
            AppendLog("[O365] EML 同步完成");
            Log.Information("[O365] EML 同步完成");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Lbl_SyncComplete"),
                LocalizationManager.GetString("Common_Complete"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            AppendLog($"[O365] EML 同步失败: {ex.Message}");
            Log.Error(ex, "[O365] EML 同步失败");
            lblO365Status.Text = LocalizationManager.GetString("Main_Lbl_O365ExportFailed");
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_SyncFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    #endregion

    #region Menu Handlers

    private void SaveSettings_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            var settings = new Dictionary<string, string>
            {
                { "txtEmlInputDir", txtEmlInputDir.Text },
                { "txtPstOutputDir", txtPstOutputDir.Text },
                { "txtEmlOutput", txtEmlOutput.Text },
                { "txtOstOutput", txtOstOutput.Text },
                { "txtImapOutputPst", txtImapOutputPst.Text },
                { "txtImapMaxEmails", txtImapMaxEmails.Text },
                { "txtSplitEmlInputDir", txtSplitEmlInputDir.Text },
                { "txtSplitBeforeDir", txtSplitBeforeDir.Text },
                { "txtSplitAfterDir", txtSplitAfterDir.Text }
            };

            var appSettingsPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "usersettings.json");
            var json = System.Text.Json.JsonSerializer.Serialize(settings);
            System.IO.File.WriteAllText(appSettingsPath, json);

            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_SettingsSaved"),
                LocalizationManager.GetString("Main_Dialog_SaveSettings"),
                MessageBoxButton.OK, MessageBoxImage.Information);
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_SettingsSaveFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    private void TabControl_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
    {
        if (e.Source is System.Windows.Controls.TabControl tabControl)
        {
            // OST 提取 tab index = 3, 邮件归类 tab index = 4
            if (tabControl.SelectedIndex == 3 || tabControl.SelectedIndex == 4)
            {
                LoadOstStoreList();
            }
        }
    }

    private void LoadSettings()
    {
        try
        {
            var appSettingsPath = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "usersettings.json");
            if (System.IO.File.Exists(appSettingsPath))
            {
                var json = System.IO.File.ReadAllText(appSettingsPath);
                var settings = System.Text.Json.JsonSerializer.Deserialize<Dictionary<string, string>>(json);

                if (settings != null)
                {
                    // EML转PST默认路径
                    if (settings.TryGetValue("DefaultEmlInput", out var defaultEmlInput) && !string.IsNullOrEmpty(defaultEmlInput)) txtEmlInputDir.Text = defaultEmlInput;
                    if (settings.TryGetValue("DefaultPstOutput", out var defaultPstOutput) && !string.IsNullOrEmpty(defaultPstOutput)) txtPstOutputDir.Text = defaultPstOutput;
                    // 其他路径
                    if (settings.TryGetValue("txtEmlOutput", out var emlOutput)) txtEmlOutput.Text = emlOutput;
                    if (settings.TryGetValue("txtOstOutput", out var ostOutput)) txtOstOutput.Text = ostOutput;
                    if (settings.TryGetValue("txtImapOutputPst", out var imapPst)) txtImapOutputPst.Text = imapPst;
                    if (settings.TryGetValue("txtImapMaxEmails", out var imapMax)) txtImapMaxEmails.Text = imapMax;
                    if (settings.TryGetValue("txtSplitEmlInputDir", out var splitInput)) txtSplitEmlInputDir.Text = splitInput;
                    if (settings.TryGetValue("txtSplitBeforeDir", out var splitBefore)) txtSplitBeforeDir.Text = splitBefore;
                    if (settings.TryGetValue("txtSplitAfterDir", out var splitAfter)) txtSplitAfterDir.Text = splitAfter;
                    // 加载IMAP默认设置
                    if (settings.TryGetValue("DefaultImapServer", out var imapServer) && !string.IsNullOrEmpty(imapServer)) txtImapServer.Text = imapServer;
                    if (settings.TryGetValue("DefaultImapPort", out var imapPort) && !string.IsNullOrEmpty(imapPort)) txtImapPort.Text = imapPort;
                    // 加载O365默认设置
                    if (settings.TryGetValue("DefaultTenantId", out var tenantId) && !string.IsNullOrEmpty(tenantId)) txtTenantId.Text = tenantId;
                    if (settings.TryGetValue("DefaultClientId", out var clientId) && !string.IsNullOrEmpty(clientId)) txtClientId.Text = clientId;
                }
            }
        }
        catch { }
    }

    private void Exit_Click(object sender, RoutedEventArgs e)
    {
        System.Windows.Application.Current.Shutdown();
    }

    private void Preferences_Click(object sender, RoutedEventArgs e)
    {
        var prefsWindow = new PreferencesWindow { Owner = this };
        prefsWindow.ShowDialog();
        LoadSettings(); // 关闭后重新加载设置
    }

    private void Activate_Click(object sender, RoutedEventArgs e)
    {
        var regWindow = new RegistrationWindow { Owner = this };
        regWindow.ShowDialog();
    }

    private void Register_Click(object sender, RoutedEventArgs e)
    {
        var regWindow = new RegistrationWindow { Owner = this };
        regWindow.ShowDialog();
    }

    private void About_Click(object sender, RoutedEventArgs e)
    {
        var settings = RegistryService.LoadRegistrationInfo();
        string licenseType = settings.IsRegistered
            ? (string.IsNullOrEmpty(settings.RegisterSerialNumber)
                ? LocalizationManager.GetString("Main_About_LicenseTrial")
                : LocalizationManager.GetString("Main_About_LicenseRegistered"))
            : LocalizationManager.GetString("Main_About_LicenseUnregistered");

        string remainingLine = settings.RegisterRemainingDays.HasValue
            ? "\n" + LocalizationManager.GetString("Main_About_RemainingDays", settings.RegisterRemainingDays.Value)
            : "";

        string aboutText =
            LocalizationManager.GetString("Main_Title") + "\n\n" +
            LocalizationManager.GetString("Main_About_Version", "1.0.0") + "\n" +
            LocalizationManager.GetString("Main_About_LicenseType", licenseType) + remainingLine + "\n\n" +
            LocalizationManager.GetString("Main_About_Features") + "\n\n" +
            LocalizationManager.GetString("Main_About_Contact", "raymond.xu@booming.one") + "\n" +
            LocalizationManager.GetString("Main_About_Copyright");

        System.Windows.MessageBox.Show(
            aboutText,
            LocalizationManager.GetString("Main_Dialog_About"),
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private void SubmitBug_Click(object sender, RoutedEventArgs e)
    {
        System.Windows.MessageBox.Show(
            LocalizationManager.GetString("Main_Msg_SubmitBug", "raymond.xu@booming.one"),
            LocalizationManager.GetString("Main_Dialog_SubmitBug"),
            MessageBoxButton.OK,
            MessageBoxImage.Information);
    }

    private void VisitWebsite_Click(object sender, RoutedEventArgs e)
    {
        try
        {
            System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
            {
                FileName = "https://www.booming.one",
                UseShellExecute = true
            });
        }
        catch (Exception ex)
        {
            System.Windows.MessageBox.Show(
                LocalizationManager.GetString("Main_Msg_OpenUrlFailed", ex.Message),
                LocalizationManager.GetString("Common_Error"),
                MessageBoxButton.OK, MessageBoxImage.Error);
        }
    }

    public void UpdateRegistrationStatus()
    {
        UpdateVersionLabel();
    }

    #endregion
}
