using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Management;
using System.Net;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

[assembly: AssemblyTitle("泽PPT备份助手")]
[assembly: AssemblyProduct("泽PPT备份助手")]
[assembly: AssemblyDescription("PPT、Word、PDF 自动备份工具")]
[assembly: AssemblyVersion("5.0.0.0")]
[assembly: AssemblyFileVersion("5.0.0.0")]
[assembly: AssemblyInformationalVersion("5.0")]

namespace ZeBackupAssistant
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            bool createdNew;
            using (Mutex mutex = new Mutex(true, "Global\\ZeBackupAssistant", out createdNew))
            {
                if (!createdNew)
                {
                    return;
                }

                using (BackupService service = new BackupService())
                {
                    service.Start();

                    if (HasArg(args, "--once"))
                    {
                        service.ScanOnce();
                        Thread.Sleep(500);
                        return;
                    }

                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new TrayContext(service));
                }
            }
        }

        private static bool HasArg(string[] args, string value)
        {
            for (int i = 0; i < args.Length; i++)
            {
                if (string.Equals(args[i], value, StringComparison.OrdinalIgnoreCase))
                {
                    return true;
                }
            }
            return false;
        }
    }

    internal sealed class TrayContext : ApplicationContext
    {
        private readonly BackupService _service;
        private readonly NotifyIcon _notifyIcon;
        private readonly Icon _appIcon;
        private MainForm _mainForm;

        public TrayContext(BackupService service)
        {
            _service = service;
            _appIcon = AppIcon.Load();
            _service.BackupSucceeded += OnBackupSucceeded;

            ContextMenuStrip menu = new ContextMenuStrip();
            ToolStripMenuItem title = new ToolStripMenuItem("泽PPT备份助手 - 运行中");
            title.Enabled = false;
            menu.Items.Add(title);
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("打开主界面", null, delegate { ShowMainForm(); });
            menu.Items.Add("立即扫描一次", null, delegate { _service.ScanOnce(); });
            menu.Items.Add("打开日志", null, delegate { _service.OpenLogFile(); });
            menu.Items.Add(new ToolStripSeparator());
            menu.Items.Add("退出", null, delegate { ExitThread(); });

            _notifyIcon = new NotifyIcon();
            _notifyIcon.Icon = _appIcon;
            _notifyIcon.Text = "泽PPT备份助手";
            _notifyIcon.ContextMenuStrip = menu;
            _notifyIcon.MouseDoubleClick += delegate { ShowMainForm(); };
            _notifyIcon.Visible = true;

            ShowDisclaimerIfNeeded();
        }

        private void OnBackupSucceeded(object sender, BackupCompletedEventArgs e)
        {
            if (!_service.SuccessTipEnabled)
            {
                return;
            }

            try
            {
                _notifyIcon.BalloonTipTitle = "备份成功";
                _notifyIcon.BalloonTipText = e.FileName + " 已完成备份。";
                _notifyIcon.ShowBalloonTip(2500);
            }
            catch
            {
            }
        }

        private void ShowMainForm()
        {
            if (_mainForm == null || _mainForm.IsDisposed)
            {
                _mainForm = new MainForm(_service, _appIcon);
            }

            _mainForm.RefreshRecords();
            _mainForm.Show();
            if (_mainForm.WindowState == FormWindowState.Minimized)
            {
                _mainForm.WindowState = FormWindowState.Normal;
            }
            _mainForm.Activate();
        }

        private void ShowDisclaimerIfNeeded()
        {
            if (_service.DisclaimerAccepted)
            {
                return;
            }

            MessageBox.Show(
                "请在本人或已授权的电脑上使用本软件。\r\n\r\n" +
                "本软件仅用于文件备份、记录查看和用户主动上传辅助。因未获授权使用、误操作、数据丢失、隐私纠纷、网络服务变化等造成的后果，由使用者自行承担，软件开发者不承担相关责任。\r\n\r\n" +
                "继续使用即表示已阅读并同意以上内容。",
                "泽PPT备份助手 免责声明",
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
            _service.AcceptDisclaimer();
        }

        protected override void ExitThreadCore()
        {
            if (_mainForm != null && !_mainForm.IsDisposed)
            {
                _mainForm.CloseForExit();
            }
            _notifyIcon.Visible = false;
            _service.BackupSucceeded -= OnBackupSucceeded;
            _notifyIcon.Dispose();
            _appIcon.Dispose();
            _service.Stop();
            base.ExitThreadCore();
        }
    }

    internal static class AppIcon
    {
        public static Icon Load()
        {
            try
            {
                Icon icon = Icon.ExtractAssociatedIcon(Application.ExecutablePath);
                if (icon != null)
                {
                    return icon;
                }
            }
            catch
            {
            }

            return (Icon)SystemIcons.Application.Clone();
        }
    }

    internal sealed class MainForm : Form
    {
        private readonly BackupService _service;
        private readonly ListView _recentList;
        private readonly Label _statusLabel;
        private readonly CheckBox _autoStartCheckBox;
        private readonly CheckBox _successTipCheckBox;
        private bool _updatingOptionControls;
        private bool _closingForExit;

        public MainForm(BackupService service, Icon appIcon)
        {
            _service = service;

            Text = "泽PPT备份助手";
            Icon = appIcon;
            StartPosition = FormStartPosition.CenterScreen;
            Size = new Size(1080, 580);
            MinimumSize = new Size(920, 500);
            ShowInTaskbar = true;
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.ColumnCount = 1;
            layout.RowCount = 3;
            layout.Padding = new Padding(14);
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 118));
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 34));
            Controls.Add(layout);

            FlowLayoutPanel toolbar = new FlowLayoutPanel();
            toolbar.Dock = DockStyle.Fill;
            toolbar.FlowDirection = FlowDirection.LeftToRight;
            toolbar.WrapContents = true;
            toolbar.Padding = new Padding(0, 0, 0, 6);
            layout.Controls.Add(toolbar, 0, 0);

            toolbar.Controls.Add(MakeButton("刷新", delegate { RefreshRecords(); }));
            toolbar.Controls.Add(MakeButton("立即扫描", delegate { _service.ScanOnce(); RefreshRecords(); }));
            toolbar.Controls.Add(MakeButton("打开日志", delegate { _service.OpenLogFile(); }));
            toolbar.Controls.Add(MakeButton("打开D盘备份", delegate { _service.OpenBackupRoot(0); }));
            toolbar.Controls.Add(MakeButton("打开E盘备份", delegate { _service.OpenBackupRoot(1); }));
            toolbar.Controls.Add(MakeButton("打开当前备份", delegate { _service.OpenCurrentBackupRoot(); }));
            toolbar.Controls.Add(MakeButton("备份位置", delegate { ShowBackupLocationSettings(); }));
            toolbar.Controls.Add(MakeButton("Wormhole网盘", delegate { OpenSelectedInWormhole(); }));
#if !NO_CLOUD
            toolbar.Controls.Add(MakeButton("上传到云盘", delegate { UploadSelectedToCloudflare(); }));
            toolbar.Controls.Add(MakeButton("打开云盘", delegate { OpenCloudflareDashboard(); }));
            toolbar.Controls.Add(MakeButton("CF说明", delegate { _service.OpenCloudflareGuide(); }));
#endif
            toolbar.Controls.Add(MakeButton("删除选中", delegate { DeleteSelectedRecord(); }));
            toolbar.Controls.Add(MakeButton("清理失效", delegate { CleanMissingRecords(); }));
            toolbar.Controls.Add(MakeButton("清空记录", delegate { ClearRecords(); }));
            toolbar.Controls.Add(MakeButton("清空日志", delegate { ClearLog(); }));
            toolbar.Controls.Add(MakeButton("自动清理", delegate { ShowAutoCleanSettings(); }));
            toolbar.Controls.Add(MakeButton("空间保护", delegate { ShowSpaceLimitSettings(); }));
            toolbar.Controls.Add(MakeButton("备份格式", delegate { ShowFormatSettings(); }));
            toolbar.Controls.Add(MakeButton("一键体检", delegate { ShowHealthCheck(); }));
            toolbar.Controls.Add(MakeButton("关于/反馈", delegate { ShowAbout(); }));

            _autoStartCheckBox = MakeCheckBox("开机自启", _service.IsAutoStartEnabled());
            _autoStartCheckBox.CheckedChanged += delegate { ToggleAutoStart(); };
            toolbar.Controls.Add(_autoStartCheckBox);

            _successTipCheckBox = MakeCheckBox("备份提示", _service.SuccessTipEnabled);
            _successTipCheckBox.CheckedChanged += delegate { ToggleSuccessTip(); };
            toolbar.Controls.Add(_successTipCheckBox);

            _recentList = new ListView();
            _recentList.Dock = DockStyle.Fill;
            _recentList.View = View.Details;
            _recentList.FullRowSelect = true;
            _recentList.GridLines = true;
            _recentList.HideSelection = false;
            _recentList.Columns.Add("时间", 145);
            _recentList.Columns.Add("类型", 80);
            _recentList.Columns.Add("文件", 170);
            _recentList.Columns.Add("大小", 80);
            _recentList.Columns.Add("状态", 70);
            _recentList.Columns.Add("备份位置", 250);
            _recentList.Columns.Add("原文件", 300);
            _recentList.DoubleClick += delegate { OpenSelectedTarget(); };
            layout.Controls.Add(_recentList, 0, 1);

            _statusLabel = new Label();
            _statusLabel.Dock = DockStyle.Fill;
            _statusLabel.TextAlign = ContentAlignment.MiddleLeft;
            _statusLabel.AutoEllipsis = true;
            layout.Controls.Add(_statusLabel, 0, 2);

            RefreshRecords();
        }

        public void RefreshRecords()
        {
            List<BackupRecord> records = _service.GetRecentBackups(5);
            _recentList.BeginUpdate();
            try
            {
                _recentList.Items.Clear();
                for (int i = 0; i < records.Count; i++)
                {
                    BackupRecord record = records[i];
                    ListViewItem item = new ListViewItem(record.TimeText);
                    item.SubItems.Add(record.Kind);
                    item.SubItems.Add(record.FileName);
                    item.SubItems.Add(record.SizeText);
                    item.SubItems.Add(record.BackupStatusText);
                    item.SubItems.Add(record.PrimaryTarget);
                    item.SubItems.Add(record.SourcePath);
                    item.Tag = record;
                    _recentList.Items.Add(item);
                }
            }
            finally
            {
                _recentList.EndUpdate();
            }

            _updatingOptionControls = true;
            try
            {
                _autoStartCheckBox.Checked = _service.IsAutoStartEnabled();
                _successTipCheckBox.Checked = _service.SuccessTipEnabled;
            }
            finally
            {
                _updatingOptionControls = false;
            }

            _statusLabel.Text = _service.GetStatusSummary(records.Count);
        }

        public void CloseForExit()
        {
            _closingForExit = true;
            Close();
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (!_closingForExit)
            {
                e.Cancel = true;
                Hide();
                return;
            }
            base.OnFormClosing(e);
        }

        private static Button MakeButton(string text, EventHandler handler)
        {
            Button button = new Button();
            button.Text = text;
            button.AutoSize = true;
            button.MinimumSize = new Size(88, 30);
            button.Margin = new Padding(0, 0, 8, 0);
            button.Click += handler;
            return button;
        }

        private static CheckBox MakeCheckBox(string text, bool isChecked)
        {
            CheckBox checkBox = new CheckBox();
            checkBox.Text = text;
            checkBox.Checked = isChecked;
            checkBox.AutoSize = true;
            checkBox.MinimumSize = new Size(86, 30);
            checkBox.Margin = new Padding(0, 4, 8, 0);
            checkBox.TextAlign = ContentAlignment.MiddleLeft;
            return checkBox;
        }

        private void ToggleAutoStart()
        {
            if (_updatingOptionControls)
            {
                return;
            }

            try
            {
                _service.SetAutoStartEnabled(_autoStartCheckBox.Checked);
                _statusLabel.Text = _autoStartCheckBox.Checked ? "已开启开机自启。" : "已关闭开机自启。";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "设置开机自启失败：\r\n\r\n" + ex.Message, "泽PPT备份助手", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                RefreshRecords();
            }
        }

        private void ToggleSuccessTip()
        {
            if (_updatingOptionControls)
            {
                return;
            }

            _service.SetSuccessTipEnabled(_successTipCheckBox.Checked);
            _statusLabel.Text = _successTipCheckBox.Checked ? "已开启备份成功提示。" : "已关闭备份成功提示。";
        }

        private void OpenSelectedTarget()
        {
            if (_recentList.SelectedItems.Count == 0)
            {
                return;
            }

            BackupRecord record = _recentList.SelectedItems[0].Tag as BackupRecord;
            if (record == null || string.IsNullOrEmpty(record.PrimaryTarget))
            {
                return;
            }

            try
            {
                if (File.Exists(record.PrimaryTarget))
                {
                    Process.Start("explorer.exe", "/select,\"" + record.PrimaryTarget + "\"");
                    return;
                }

                string directory = Path.GetDirectoryName(record.PrimaryTarget);
                if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory))
                {
                    Process.Start("explorer.exe", "\"" + directory + "\"");
                }
            }
            catch
            {
            }
        }

        private BackupRecord GetSelectedOrNewestRecord()
        {
            if (_recentList.SelectedItems.Count > 0)
            {
                return _recentList.SelectedItems[0].Tag as BackupRecord;
            }

            if (_recentList.Items.Count > 0)
            {
                return _recentList.Items[0].Tag as BackupRecord;
            }

            MessageBox.Show(this, "还没有备份记录，先打开一个 PPT / Word / PDF 等它备份一次。", "泽PPT备份助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return null;
        }

        private string GetExistingBackupPathOrShow(BackupRecord record)
        {
            string path = _service.FindExistingBackupFile(record);
            if (!string.IsNullOrEmpty(path))
            {
                return path;
            }

            MessageBox.Show(this, "这条记录对应的 D/E 备份文件已经不存在了，可以点“清理失效”后重新备份。", "泽PPT备份助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
            return "";
        }

        private void OpenSelectedInWormhole()
        {
            BackupRecord record = GetSelectedOrNewestRecord();
            if (record == null)
            {
                return;
            }

            string path = GetExistingBackupPathOrShow(record);
            if (string.IsNullOrEmpty(path))
            {
                return;
            }

            _service.OpenWormholeForFile(path);
            _statusLabel.Text = "已打开 Wormhole网盘，并选中/复制文件路径：" + Path.GetFileName(path);
        }

        private void UploadSelectedToCloudflare()
        {
            BackupRecord record = GetSelectedOrNewestRecord();
            if (record == null)
            {
                return;
            }

            RunUpload("Cloudflare", record, delegate(BackupRecord uploadRecord)
            {
                return _service.UploadToCloudflare(uploadRecord);
            });
        }

        private void OpenCloudflareDashboard()
        {
            try
            {
                try
                {
                    Clipboard.SetText(_service.GetCloudflareAdminToken());
                }
                catch
                {
                }
                Process.Start(_service.GetCloudflareDashboardUrl());
                _statusLabel.Text = "已打开云盘后台，管理口令已复制。";
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, "打开云盘失败：\r\n\r\n" + ex.Message, "泽PPT备份助手", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private void RunUpload(string serviceName, BackupRecord record, Func<BackupRecord, string> uploadAction)
        {
            string path = GetExistingBackupPathOrShow(record);
            if (string.IsNullOrEmpty(path))
            {
                return;
            }

            _statusLabel.Text = "正在上传到 " + serviceName + "：" + Path.GetFileName(path);

            ThreadPool.QueueUserWorkItem(delegate
            {
                try
                {
                    string link = uploadAction(record);
                    BeginInvoke((MethodInvoker)delegate
                    {
                        try
                        {
                            Clipboard.SetText(link);
                        }
                        catch
                        {
                        }

                        _statusLabel.Text = serviceName + " 上传成功，链接已复制，也可以在云盘后台查看。";
                        MessageBox.Show(this, "上传成功，链接已经复制到剪贴板。\r\n\r\n以后也可以点“打开云盘”查看上传记录。\r\n\r\n" + link, "泽PPT备份助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    });
                }
                catch (Exception ex)
                {
                    BeginInvoke((MethodInvoker)delegate
                    {
                        _statusLabel.Text = serviceName + " 上传失败：" + ex.Message;
                        MessageBox.Show(this, serviceName + " 上传失败：\r\n\r\n" + ex.Message, "泽PPT备份助手", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    });
                }
            });
        }

        private void DeleteSelectedRecord()
        {
            if (_recentList.SelectedItems.Count == 0)
            {
                MessageBox.Show(this, "先在列表里选中一条备份记录。", "泽PPT备份助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            BackupRecord record = _recentList.SelectedItems[0].Tag as BackupRecord;
            if (record == null)
            {
                return;
            }

            DialogResult result = MessageBox.Show(
                this,
                "删除选中记录会移除去重记录，并尝试删除对应的备份文件。\r\n\r\n源文件不会被删除。如果源文件当前还打开，本次运行内不会马上重新备份同一个版本。继续吗？",
                "泽PPT备份助手",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (result != DialogResult.Yes)
            {
                return;
            }

            int deletedFiles = _service.DeleteBackupRecord(record, true);
            RefreshRecords();
            MessageBox.Show(this, "已删除记录。删除的备份文件数量：" + deletedFiles.ToString() + "\r\n\r\n如果源文件仍然打开，程序不会在本次运行内马上重新备份同一个版本。", "泽PPT备份助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void CleanMissingRecords()
        {
            int removed = _service.CleanMissingBackupRecords();
            RefreshRecords();
            MessageBox.Show(this, "已清理失效记录：" + removed.ToString() + " 条。", "泽PPT备份助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ClearRecords()
        {
            DialogResult result = MessageBox.Show(
                this,
                "清空记录会清空去重索引，之后已经打开过的文件也可以重新备份。\r\n\r\n不会删除源文件，也不会删除 D/E 盘已经存在的备份文件。继续吗？",
                "泽PPT备份助手",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (result != DialogResult.Yes)
            {
                return;
            }

            _service.ClearBackupRecords();
            RefreshRecords();
            MessageBox.Show(this, "备份记录已清空。", "泽PPT备份助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ClearLog()
        {
            DialogResult result = MessageBox.Show(
                this,
                "清空日志只会清空 D 盘日志文本，不影响备份文件和去重记录。继续吗？",
                "泽PPT备份助手",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Question
            );

            if (result != DialogResult.Yes)
            {
                return;
            }

            _service.ClearLogFile();
            MessageBox.Show(this, "日志已清空。", "泽PPT备份助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void ShowAutoCleanSettings()
        {
            using (AutoCleanDialog dialog = new AutoCleanDialog(_service.AutoCleanDays))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                _service.SetAutoCleanDays(dialog.SelectedDays);
                int removed = _service.RunAutoCleanNow();
                RefreshRecords();

                string message;
                if (dialog.SelectedDays <= 0)
                {
                    message = "已设置为永久保留备份文件。";
                }
                else
                {
                    message = "已设置为保留最近 " + dialog.SelectedDays.ToString() + " 天备份。\r\n\r\n本次清理旧日期文件夹：" + removed.ToString() + " 个。";
                }
                MessageBox.Show(this, message, "泽PPT备份助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ShowSpaceLimitSettings()
        {
            using (SpaceLimitDialog dialog = new SpaceLimitDialog(_service.MaxBackupSizeMb))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                _service.SetMaxBackupSizeMb(dialog.SelectedMegabytes);
                CleanupResult result = _service.RunSpaceLimitCleanNow();
                RefreshRecords();

                string message = dialog.SelectedMegabytes <= 0
                    ? "已关闭磁盘空间保护。"
                    : "已设置为每个备份位置最多占用 " + BackupService.FormatSize(dialog.SelectedMegabytes * 1024L * 1024L) + "。";

                message += "\r\n\r\n本次清理文件：" + result.DeletedFiles.ToString() + " 个，释放空间：" + BackupService.FormatSize(result.FreedBytes) + "。";
                MessageBox.Show(this, message, "泽PPT备份助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ShowFormatSettings()
        {
            using (FormatSettingsDialog dialog = new FormatSettingsDialog(_service.BackupPptEnabled, _service.BackupWordEnabled, _service.BackupPdfEnabled))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                _service.SetBackupFormatOptions(dialog.PptEnabled, dialog.WordEnabled, dialog.PdfEnabled);
                RefreshRecords();
                MessageBox.Show(this, "已更新备份格式：\r\n\r\n" + _service.BackupFormatDescription, "泽PPT备份助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ShowBackupLocationSettings()
        {
            using (BackupLocationDialog dialog = new BackupLocationDialog(_service.CustomBackupRoot))
            {
                if (dialog.ShowDialog(this) != DialogResult.OK)
                {
                    return;
                }

                _service.SetCustomBackupRoot(dialog.SelectedPath);
                RefreshRecords();
                MessageBox.Show(this, "备份位置已更新：\r\n\r\n" + _service.BackupLocationDescription, "泽PPT备份助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
        }

        private void ShowHealthCheck()
        {
            using (TextReportDialog dialog = new TextReportDialog("一键体检", _service.BuildHealthReport()))
            {
                dialog.ShowDialog(this);
            }
        }

        private void ShowAbout()
        {
            using (AboutDialog dialog = new AboutDialog())
            {
                dialog.ShowDialog(this);
            }
        }
    }

    internal sealed class AboutDialog : Form
    {
        private const string FeedbackEmail = "zeningshuyi@gmail.com";
        private readonly Label _copyStatusLabel;

        public AboutDialog()
        {
            Text = "关于/反馈";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(470, 260);
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            Label title = new Label();
            title.Text = "泽PPT备份助手";
            title.Font = new Font(Font.FontFamily, 14F, FontStyle.Bold, GraphicsUnit.Point);
            title.AutoSize = true;
            title.Location = new Point(22, 20);
            Controls.Add(title);

            Label version = new Label();
            version.Text = "当前版本：v" + GetVersionText();
            version.AutoSize = true;
            version.Location = new Point(24, 58);
            Controls.Add(version);

            Label description = new Label();
            description.Text = "支持 PPT / Word / PDF 自动备份，支持按日期分类、重名编号、去重、日志、最近记录、空间保护和自定义备份设置。";
            description.Location = new Point(24, 88);
            description.Size = new Size(420, 42);
            Controls.Add(description);

            Label emailLabel = new Label();
            emailLabel.Text = "反馈邮箱：";
            emailLabel.AutoSize = true;
            emailLabel.Location = new Point(24, 142);
            Controls.Add(emailLabel);

            TextBox emailBox = new TextBox();
            emailBox.Text = FeedbackEmail;
            emailBox.ReadOnly = true;
            emailBox.Location = new Point(96, 138);
            emailBox.Width = 220;
            Controls.Add(emailBox);

            Button copyEmail = new Button();
            copyEmail.Text = "复制邮箱";
            copyEmail.Location = new Point(326, 136);
            copyEmail.Size = new Size(90, 28);
            copyEmail.Click += delegate { CopyEmail(); };
            Controls.Add(copyEmail);

            _copyStatusLabel = new Label();
            _copyStatusLabel.Text = "";
            _copyStatusLabel.AutoSize = true;
            _copyStatusLabel.Location = new Point(96, 170);
            Controls.Add(_copyStatusLabel);

            Button close = new Button();
            close.Text = "关闭";
            close.DialogResult = DialogResult.OK;
            close.Location = new Point(342, 212);
            close.Size = new Size(75, 28);
            Controls.Add(close);

            AcceptButton = close;
        }

        private void CopyEmail()
        {
            try
            {
                Clipboard.SetText(FeedbackEmail);
                _copyStatusLabel.Text = "邮箱已复制。";
            }
            catch (Exception ex)
            {
                _copyStatusLabel.Text = "复制失败：" + ex.Message;
            }
        }

        private static string GetVersionText()
        {
            try
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyInformationalVersionAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyInformationalVersionAttribute info = attributes[0] as AssemblyInformationalVersionAttribute;
                    if (info != null && !string.IsNullOrEmpty(info.InformationalVersion))
                    {
                        return info.InformationalVersion;
                    }
                }
            }
            catch
            {
            }

            return "5.0";
        }
    }

    internal sealed class AutoCleanDialog : Form
    {
        private readonly ComboBox _comboBox;

        public int SelectedDays { get; private set; }

        public AutoCleanDialog(int currentDays)
        {
            Text = "自动清理设置";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(360, 150);
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            Label label = new Label();
            label.Text = "选择备份保留时间：";
            label.AutoSize = true;
            label.Location = new Point(18, 18);
            Controls.Add(label);

            _comboBox = new ComboBox();
            _comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            _comboBox.Location = new Point(20, 48);
            _comboBox.Width = 310;
            _comboBox.Items.Add(new RetentionOption("永久保留，不自动删除", 0));
            _comboBox.Items.Add(new RetentionOption("保留最近 30 天", 30));
            _comboBox.Items.Add(new RetentionOption("保留最近 60 天", 60));
            _comboBox.Items.Add(new RetentionOption("保留最近 90 天", 90));
            Controls.Add(_comboBox);

            Button ok = new Button();
            ok.Text = "确定";
            ok.DialogResult = DialogResult.OK;
            ok.Location = new Point(172, 102);
            ok.Size = new Size(75, 28);
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "取消";
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Location = new Point(255, 102);
            cancel.Size = new Size(75, 28);
            Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;

            SelectCurrentOption(currentDays);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
            {
                RetentionOption option = _comboBox.SelectedItem as RetentionOption;
                if (option != null)
                {
                    SelectedDays = option.Days;
                }
            }
            base.OnFormClosing(e);
        }

        private void SelectCurrentOption(int days)
        {
            for (int i = 0; i < _comboBox.Items.Count; i++)
            {
                RetentionOption option = _comboBox.Items[i] as RetentionOption;
                if (option != null && option.Days == days)
                {
                    _comboBox.SelectedIndex = i;
                    return;
                }
            }
            _comboBox.SelectedIndex = 0;
        }

        private sealed class RetentionOption
        {
            public readonly string Text;
            public readonly int Days;

            public RetentionOption(string text, int days)
            {
                Text = text;
                Days = days;
            }

            public override string ToString()
            {
                return Text;
            }
        }
    }

    internal sealed class SpaceLimitDialog : Form
    {
        private readonly ComboBox _comboBox;

        public int SelectedMegabytes { get; private set; }

        public SpaceLimitDialog(int currentMegabytes)
        {
            Text = "磁盘空间保护";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(390, 160);
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            Label label = new Label();
            label.Text = "选择每个备份位置最多占用空间：";
            label.AutoSize = true;
            label.Location = new Point(18, 18);
            Controls.Add(label);

            _comboBox = new ComboBox();
            _comboBox.DropDownStyle = ComboBoxStyle.DropDownList;
            _comboBox.Location = new Point(20, 50);
            _comboBox.Width = 340;
            _comboBox.Items.Add(new SpaceLimitOption("不限制", 0));
            _comboBox.Items.Add(new SpaceLimitOption("最多 2 GB", 2048));
            _comboBox.Items.Add(new SpaceLimitOption("最多 5 GB", 5120));
            _comboBox.Items.Add(new SpaceLimitOption("最多 10 GB", 10240));
            _comboBox.Items.Add(new SpaceLimitOption("最多 20 GB", 20480));
            _comboBox.Items.Add(new SpaceLimitOption("最多 50 GB", 51200));
            Controls.Add(_comboBox);

            Button ok = new Button();
            ok.Text = "确定";
            ok.DialogResult = DialogResult.OK;
            ok.Location = new Point(202, 106);
            ok.Size = new Size(75, 28);
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "取消";
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Location = new Point(285, 106);
            cancel.Size = new Size(75, 28);
            Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;
            SelectCurrentOption(currentMegabytes);
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
            {
                SpaceLimitOption option = _comboBox.SelectedItem as SpaceLimitOption;
                if (option != null)
                {
                    SelectedMegabytes = option.Megabytes;
                }
            }
            base.OnFormClosing(e);
        }

        private void SelectCurrentOption(int megabytes)
        {
            for (int i = 0; i < _comboBox.Items.Count; i++)
            {
                SpaceLimitOption option = _comboBox.Items[i] as SpaceLimitOption;
                if (option != null && option.Megabytes == megabytes)
                {
                    _comboBox.SelectedIndex = i;
                    return;
                }
            }
            _comboBox.SelectedIndex = 0;
        }

        private sealed class SpaceLimitOption
        {
            public readonly string Text;
            public readonly int Megabytes;

            public SpaceLimitOption(string text, int megabytes)
            {
                Text = text;
                Megabytes = megabytes;
            }

            public override string ToString()
            {
                return Text;
            }
        }
    }

    internal sealed class FormatSettingsDialog : Form
    {
        private readonly CheckBox _pptCheckBox;
        private readonly CheckBox _wordCheckBox;
        private readonly CheckBox _pdfCheckBox;

        public bool PptEnabled { get; private set; }
        public bool WordEnabled { get; private set; }
        public bool PdfEnabled { get; private set; }

        public FormatSettingsDialog(bool pptEnabled, bool wordEnabled, bool pdfEnabled)
        {
            Text = "备份格式";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(360, 205);
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            Label label = new Label();
            label.Text = "选择需要自动备份的文件类型：";
            label.AutoSize = true;
            label.Location = new Point(18, 18);
            Controls.Add(label);

            _pptCheckBox = MakeFormatCheckBox("PPT / PPTX / PPS / PPSX", pptEnabled, 20, 52);
            _wordCheckBox = MakeFormatCheckBox("Word / WPS / RTF", wordEnabled, 20, 82);
            _pdfCheckBox = MakeFormatCheckBox("PDF", pdfEnabled, 20, 112);
            Controls.Add(_pptCheckBox);
            Controls.Add(_wordCheckBox);
            Controls.Add(_pdfCheckBox);

            Button ok = new Button();
            ok.Text = "确定";
            ok.DialogResult = DialogResult.OK;
            ok.Location = new Point(172, 156);
            ok.Size = new Size(75, 28);
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "取消";
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Location = new Point(255, 156);
            cancel.Size = new Size(75, 28);
            Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
            {
                if (!_pptCheckBox.Checked && !_wordCheckBox.Checked && !_pdfCheckBox.Checked)
                {
                    MessageBox.Show(this, "至少要保留一种备份格式。", "泽PPT备份助手", MessageBoxButtons.OK, MessageBoxIcon.Information);
                    e.Cancel = true;
                    return;
                }

                PptEnabled = _pptCheckBox.Checked;
                WordEnabled = _wordCheckBox.Checked;
                PdfEnabled = _pdfCheckBox.Checked;
            }
            base.OnFormClosing(e);
        }

        private static CheckBox MakeFormatCheckBox(string text, bool isChecked, int x, int y)
        {
            CheckBox checkBox = new CheckBox();
            checkBox.Text = text;
            checkBox.Checked = isChecked;
            checkBox.AutoSize = true;
            checkBox.Location = new Point(x, y);
            return checkBox;
        }
    }

    internal sealed class BackupLocationDialog : Form
    {
        private readonly TextBox _pathBox;

        public string SelectedPath { get; private set; }

        public BackupLocationDialog(string currentPath)
        {
            Text = "备份位置";
            StartPosition = FormStartPosition.CenterParent;
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            ClientSize = new Size(560, 190);
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            Label label = new Label();
            label.Text = "自定义备份目录。留空表示使用默认 D/E/C 自动备份位置。";
            label.AutoSize = true;
            label.Location = new Point(18, 18);
            Controls.Add(label);

            _pathBox = new TextBox();
            _pathBox.Location = new Point(20, 52);
            _pathBox.Width = 420;
            _pathBox.Text = currentPath ?? "";
            Controls.Add(_pathBox);

            Button browse = new Button();
            browse.Text = "选择";
            browse.Location = new Point(452, 50);
            browse.Size = new Size(75, 28);
            browse.Click += delegate { BrowseFolder(); };
            Controls.Add(browse);

            Button useDefault = new Button();
            useDefault.Text = "使用默认";
            useDefault.Location = new Point(20, 102);
            useDefault.Size = new Size(95, 28);
            useDefault.Click += delegate { _pathBox.Text = ""; };
            Controls.Add(useDefault);

            Button ok = new Button();
            ok.Text = "确定";
            ok.DialogResult = DialogResult.OK;
            ok.Location = new Point(370, 132);
            ok.Size = new Size(75, 28);
            Controls.Add(ok);

            Button cancel = new Button();
            cancel.Text = "取消";
            cancel.DialogResult = DialogResult.Cancel;
            cancel.Location = new Point(452, 132);
            cancel.Size = new Size(75, 28);
            Controls.Add(cancel);

            AcceptButton = ok;
            CancelButton = cancel;
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            if (DialogResult == DialogResult.OK)
            {
                string value = (_pathBox.Text ?? "").Trim();
                if (value.Length == 0)
                {
                    SelectedPath = "";
                }
                else
                {
                    try
                    {
                        SelectedPath = Path.GetFullPath(value);
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(this, "备份目录路径无效：\r\n\r\n" + ex.Message, "泽PPT备份助手", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        e.Cancel = true;
                        return;
                    }
                }
            }
            base.OnFormClosing(e);
        }

        private void BrowseFolder()
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.Description = "选择泽PPT备份助手的备份保存目录";
                dialog.ShowNewFolderButton = true;
                if (Directory.Exists(_pathBox.Text))
                {
                    dialog.SelectedPath = _pathBox.Text;
                }

                if (dialog.ShowDialog(this) == DialogResult.OK)
                {
                    _pathBox.Text = dialog.SelectedPath;
                }
            }
        }
    }

    internal sealed class TextReportDialog : Form
    {
        private readonly TextBox _textBox;

        public TextReportDialog(string title, string report)
        {
            Text = "泽PPT备份助手 - " + title;
            StartPosition = FormStartPosition.CenterParent;
            Size = new Size(760, 560);
            MinimumSize = new Size(620, 420);
            ShowInTaskbar = false;
            Font = new Font("Microsoft YaHei UI", 9F, FontStyle.Regular, GraphicsUnit.Point);

            TableLayoutPanel layout = new TableLayoutPanel();
            layout.Dock = DockStyle.Fill;
            layout.ColumnCount = 1;
            layout.RowCount = 2;
            layout.Padding = new Padding(12);
            layout.RowStyles.Add(new RowStyle(SizeType.Percent, 100));
            layout.RowStyles.Add(new RowStyle(SizeType.Absolute, 42));
            Controls.Add(layout);

            _textBox = new TextBox();
            _textBox.Dock = DockStyle.Fill;
            _textBox.Multiline = true;
            _textBox.ReadOnly = true;
            _textBox.ScrollBars = ScrollBars.Both;
            _textBox.WordWrap = false;
            _textBox.Text = report ?? "";
            layout.Controls.Add(_textBox, 0, 0);

            FlowLayoutPanel buttons = new FlowLayoutPanel();
            buttons.Dock = DockStyle.Fill;
            buttons.FlowDirection = FlowDirection.RightToLeft;
            layout.Controls.Add(buttons, 0, 1);

            Button close = new Button();
            close.Text = "关闭";
            close.DialogResult = DialogResult.OK;
            close.Size = new Size(75, 28);
            buttons.Controls.Add(close);

            Button copy = new Button();
            copy.Text = "复制";
            copy.Size = new Size(75, 28);
            copy.Click += delegate
            {
                try
                {
                    Clipboard.SetText(_textBox.Text);
                }
                catch
                {
                }
            };
            buttons.Controls.Add(copy);

            AcceptButton = close;
        }
    }

    internal sealed class BackupService : IDisposable
    {
        private const string BackupFolderName = "泽宁PPPPPPPPTTTT备份";
        private static readonly string[] BackupRoots = new string[]
        {
            @"D:\" + BackupFolderName,
            @"E:\" + BackupFolderName
        };
        private const string WormholeUrl = "https://wormhole.app/";
        private const string UploadConfigFileName = "upload-config.ini";
        private const string SettingsFileName = "settings.ini";
        private const int DefaultMaxBackupSizeMb = 0;
        private const string AutoStartValueName = "ZePPTBackupAssistant";
        private const string RunRegistryPath = @"Software\Microsoft\Windows\CurrentVersion\Run";
        private static string s_customBackupRoot = "";

        private static readonly string[] PresentationProgIds = new string[]
        {
            "PowerPoint.Application",
            "KWPP.Application",
            "KWPP.Application.12",
            "KWPP.Application.11",
            "WPP.Application",
            "WPP.Application.12"
        };

        private static readonly string[] WordProgIds = new string[]
        {
            "Word.Application",
            "KWps.Application",
            "KWps.Application.12",
            "KWps.Application.11",
            "Wps.Application"
        };

        private static readonly string[] AllowedExtensions = new string[]
        {
            ".ppt",
            ".pptx",
            ".pps",
            ".ppsx",
            ".doc",
            ".docx",
            ".docm",
            ".rtf",
            ".wps",
            ".pdf"
        };

        private readonly object _gate = new object();
        private readonly Dictionary<string, PendingDocument> _pending = new Dictionary<string, PendingDocument>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _backedUpFingerprints = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _fingerprintSavedPaths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, string> _contentHashSavedPaths = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, RecentNameSizeEntry> _recentNameSizeEntries = new Dictionary<string, RecentNameSizeEntry>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _suppressedFingerprints = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _rootStatusLogged = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly HashSet<string> _progIdErrorLogged = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        private readonly Dictionary<string, TitleSearchCacheEntry> _pdfTitleCache = new Dictionary<string, TitleSearchCacheEntry>(StringComparer.OrdinalIgnoreCase);
        private string _dataRoot;
        private System.Threading.Timer _timer;
        private bool _scanning;
        private bool _successTipEnabled;
        private bool _backupPptEnabled;
        private bool _backupWordEnabled;
        private bool _backupPdfEnabled;
        private bool _disclaimerAccepted;
        private string _customBackupRoot;
        private int _autoCleanDays;
        private int _maxBackupSizeMb;
        private DateTime _lastAutoCleanDate = DateTime.MinValue;
        private DateTime _lastScanTime = DateTime.MinValue;
        private int _lastScanDocumentCount;
        private DateTime _lastBackupTime = DateTime.MinValue;
        private string _lastBackupFileName = "";

        public event EventHandler<BackupCompletedEventArgs> BackupSucceeded;

        public string LogFilePath
        {
            get { return Path.Combine(GetDataRoot(), "日志", "泽.log"); }
        }

        public string IndexFilePath
        {
            get { return Path.Combine(GetDataRoot(), "日志", "backup-index.tsv"); }
        }

        public string UploadConfigPath
        {
            get { return Path.Combine(GetDataRoot(), "日志", UploadConfigFileName); }
        }

        public string SettingsFilePath
        {
            get { return Path.Combine(GetDataRoot(), "日志", SettingsFileName); }
        }

        public bool SuccessTipEnabled
        {
            get { return _successTipEnabled; }
        }

        public bool BackupPptEnabled
        {
            get { return _backupPptEnabled; }
        }

        public bool BackupWordEnabled
        {
            get { return _backupWordEnabled; }
        }

        public bool BackupPdfEnabled
        {
            get { return _backupPdfEnabled; }
        }

        public bool DisclaimerAccepted
        {
            get { return _disclaimerAccepted; }
        }

        public string CustomBackupRoot
        {
            get { return _customBackupRoot ?? ""; }
        }

        public int AutoCleanDays
        {
            get { return _autoCleanDays; }
        }

        public int MaxBackupSizeMb
        {
            get { return _maxBackupSizeMb; }
        }

        public string AutoCleanDescription
        {
            get
            {
                if (_autoCleanDays <= 0)
                {
                    return "永久保留";
                }
                return "保留 " + _autoCleanDays.ToString() + " 天";
            }
        }

        public string SpaceLimitDescription
        {
            get
            {
                if (_maxBackupSizeMb <= 0)
                {
                    return "不限制";
                }
                return "每个位置最多 " + FormatSize(_maxBackupSizeMb * 1024L * 1024L);
            }
        }

        public string BackupFormatDescription
        {
            get
            {
                List<string> parts = new List<string>();
                if (_backupPptEnabled)
                {
                    parts.Add("PPT");
                }
                if (_backupWordEnabled)
                {
                    parts.Add("Word");
                }
                if (_backupPdfEnabled)
                {
                    parts.Add("PDF");
                }
                return parts.Count == 0 ? "未启用" : string.Join(" / ", parts.ToArray());
            }
        }

        public string BackupLocationDescription
        {
            get
            {
                if (!string.IsNullOrWhiteSpace(_customBackupRoot))
                {
                    return "自定义：" + _customBackupRoot;
                }
                return "默认：D/E 优先，不可用时自动使用 C 盘或文档目录";
            }
        }

        public void Start()
        {
            LoadSettings();
            EnsureLogDirectory();
            LoadIndex();
            EnsureBackupRoots();
            CleanOldBackupsIfNeeded(false);
            RunSpaceLimitCleanNow();
            Log("程序启动。静默检测：" + BackupFormatDescription + "。");
            _timer = new System.Threading.Timer(delegate { ScanOnce(); }, null, 1000, 3000);
        }

        public void Stop()
        {
            System.Threading.Timer timer = _timer;
            _timer = null;
            if (timer != null)
            {
                timer.Dispose();
            }
            Log("程序退出。");
        }

        public void Dispose()
        {
            Stop();
        }

        public void ScanOnce()
        {
            lock (_gate)
            {
                if (_scanning)
                {
                    return;
                }
                _scanning = true;
            }

            try
            {
                EnsureBackupRoots();

                List<OpenDocument> opened = new List<OpenDocument>();
                if (_backupPptEnabled)
                {
                    opened.AddRange(GetOpenPresentations());
                }
                if (_backupWordEnabled)
                {
                    opened.AddRange(GetOpenWordDocuments());
                }
                if (_backupPdfEnabled)
                {
                    opened.AddRange(GetOpenPdfDocuments());
                }

                _lastScanTime = DateTime.Now;
                _lastScanDocumentCount = opened.Count;

                DateTime now = DateTime.UtcNow;
                for (int i = 0; i < opened.Count; i++)
                {
                    HandleDocument(opened[i], now);
                }
            }
            catch (Exception ex)
            {
                Log("扫描异常：" + ex.Message);
            }
            finally
            {
                lock (_gate)
                {
                    _scanning = false;
                }
            }
        }

        public void OpenLogFile()
        {
            try
            {
                if (!File.Exists(LogFilePath))
                {
                    Directory.CreateDirectory(Path.GetDirectoryName(LogFilePath));
                    File.WriteAllText(LogFilePath, "", Encoding.UTF8);
                }
                Process.Start("notepad.exe", LogFilePath);
            }
            catch
            {
            }
        }

        public void OpenBackupRoot(int index)
        {
            try
            {
                if (index < 0 || index >= BackupRoots.Length)
                {
                    return;
                }

                string root = BackupRoots[index];
                if (EnsureBackupRoot(root))
                {
                    Process.Start("explorer.exe", "\"" + root + "\"");
                }
            }
            catch
            {
            }
        }

        public void OpenCurrentBackupRoot()
        {
            try
            {
                List<string> roots = GetActiveBackupRoots();
                if (roots.Count > 0)
                {
                    Process.Start("explorer.exe", "\"" + roots[0] + "\"");
                }
            }
            catch
            {
            }
        }

        public bool IsAutoStartEnabled()
        {
            try
            {
                using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(RunRegistryPath, false))
                {
                    if (key == null)
                    {
                        return false;
                    }

                    object value = key.GetValue(AutoStartValueName);
                    if (value == null)
                    {
                        return false;
                    }

                    string text = Convert.ToString(value);
                    return !string.IsNullOrWhiteSpace(text);
                }
            }
            catch
            {
                return false;
            }
        }

        public void SetAutoStartEnabled(bool enabled)
        {
            using (Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(RunRegistryPath))
            {
                if (key == null)
                {
                    throw new InvalidOperationException("无法打开当前用户启动项。");
                }

                if (enabled)
                {
                    key.SetValue(AutoStartValueName, "\"" + Application.ExecutablePath + "\"");
                    Log("已开启开机自启：" + Application.ExecutablePath);
                }
                else
                {
                    key.DeleteValue(AutoStartValueName, false);
                    Log("已关闭开机自启。");
                }
            }
        }

        public void SetSuccessTipEnabled(bool enabled)
        {
            _successTipEnabled = enabled;
            SaveSettings();
            Log(enabled ? "已开启备份成功提示。" : "已关闭备份成功提示。");
        }

        public void AcceptDisclaimer()
        {
            _disclaimerAccepted = true;
            SaveSettings();
            Log("用户已确认免责声明。");
        }

        public void SetBackupFormatOptions(bool pptEnabled, bool wordEnabled, bool pdfEnabled)
        {
            if (!pptEnabled && !wordEnabled && !pdfEnabled)
            {
                pptEnabled = true;
            }

            _backupPptEnabled = pptEnabled;
            _backupWordEnabled = wordEnabled;
            _backupPdfEnabled = pdfEnabled;
            SaveSettings();
            Log("备份格式已更新：" + BackupFormatDescription + "。");
        }

        public void SetCustomBackupRoot(string path)
        {
            string normalized = "";
            if (!string.IsNullOrWhiteSpace(path))
            {
                normalized = Path.GetFullPath(path.Trim());
            }

            _customBackupRoot = normalized;
            s_customBackupRoot = normalized;
            _dataRoot = null;
            SaveSettings();
            EnsureBackupRoots();
            LoadIndex();
            Log(string.IsNullOrEmpty(normalized) ? "备份位置已恢复默认。" : "备份位置已设置为自定义目录：" + normalized);
        }

        public void SetAutoCleanDays(int days)
        {
            if (days != 0 && days != 30 && days != 60 && days != 90)
            {
                days = 0;
            }

            _autoCleanDays = days;
            _lastAutoCleanDate = DateTime.MinValue;
            SaveSettings();
            Log(_autoCleanDays <= 0 ? "自动清理已设置为永久保留。" : "自动清理已设置为保留最近 " + _autoCleanDays.ToString() + " 天。");
        }

        public void SetMaxBackupSizeMb(int megabytes)
        {
            if (megabytes != 0 && megabytes != 2048 && megabytes != 5120 && megabytes != 10240 && megabytes != 20480 && megabytes != 51200)
            {
                megabytes = DefaultMaxBackupSizeMb;
            }

            _maxBackupSizeMb = megabytes;
            SaveSettings();
            Log(_maxBackupSizeMb <= 0 ? "磁盘空间保护已关闭。" : "磁盘空间保护已设置为：" + SpaceLimitDescription + "。");
        }

        public int RunAutoCleanNow()
        {
            int removed = CleanOldBackupsIfNeeded(true);
            if (removed > 0)
            {
                CleanMissingBackupRecords();
            }
            return removed;
        }

        public CleanupResult RunSpaceLimitCleanNow()
        {
            CleanupResult result = CleanBackupsBySizeLimit();
            if (result.DeletedFiles > 0)
            {
                CleanMissingBackupRecords();
            }
            return result;
        }

        public string FindExistingBackupFile(BackupRecord record)
        {
            if (record == null)
            {
                return "";
            }

            string[] paths = SplitSavedPaths(record.SavedPaths);
            for (int i = 0; i < paths.Length; i++)
            {
                string path = paths[i];
                if (!string.IsNullOrWhiteSpace(path) && File.Exists(path))
                {
                    return path;
                }
            }

            if (!string.IsNullOrWhiteSpace(record.PrimaryTarget) && File.Exists(record.PrimaryTarget))
            {
                return record.PrimaryTarget;
            }

            return "";
        }

        public void OpenWormholeForFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
            {
                throw new FileNotFoundException("没有找到要上传的备份文件。", path);
            }

            try
            {
                Clipboard.SetText(path);
            }
            catch
            {
            }

            try
            {
                Process.Start("explorer.exe", "/select,\"" + path + "\"");
            }
            catch
            {
            }

            Process.Start(WormholeUrl);
            Log("已打开 Wormhole网盘，并选中文件：" + path);
        }

        public void OpenCloudflareGuide()
        {
            try
            {
                string projectPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "cloudflare-pages-upload");
                string guidePath = Path.Combine(projectPath, "使用说明-CloudflarePages上传页.txt");
                if (File.Exists(guidePath))
                {
                    Process.Start("notepad.exe", guidePath);
                    return;
                }

                if (Directory.Exists(projectPath))
                {
                    Process.Start("explorer.exe", "\"" + projectPath + "\"");
                }
            }
            catch
            {
            }
        }

        public string UploadToCloudflare(BackupRecord record)
        {
            string path = FindExistingBackupFile(record);
            if (string.IsNullOrEmpty(path))
            {
                throw new FileNotFoundException("这条记录对应的 D/E 备份文件已经不存在。");
            }

            UploadConfig config = LoadUploadConfig();
            string link = UploadRawToCloudflare(path, config);
            Log("Cloudflare 上传成功：" + path + " -> " + link);
            return link;
        }

        public string GetCloudflareDashboardUrl()
        {
            UploadConfig config = LoadUploadConfig();
            string url = config.CloudflareUploadUrl;
            if (url.EndsWith("/api/upload", StringComparison.OrdinalIgnoreCase))
            {
                url = url.Substring(0, url.Length - "/api/upload".Length);
            }
            return url.TrimEnd('/') + "/admin.html#token=" + Uri.EscapeDataString(config.AdminToken);
        }

        public string GetCloudflareAdminToken()
        {
            return LoadUploadConfig().AdminToken;
        }

        public List<BackupRecord> GetRecentBackups(int count)
        {
            List<BackupRecord> records = new List<BackupRecord>();

            try
            {
                if (!File.Exists(IndexFilePath))
                {
                    return records;
                }

                string[] lines = File.ReadAllLines(IndexFilePath, Encoding.UTF8);
                for (int i = lines.Length - 1; i >= 0 && records.Count < count; i--)
                {
                    BackupRecord record = ParseBackupRecord(lines[i]);
                    if (record != null)
                    {
                        records.Add(record);
                    }
                }
            }
            catch
            {
            }

            return records;
        }

        public string GetStatusSummary(int recentCount)
        {
            string scanText = _lastScanTime == DateTime.MinValue ? "尚未扫描" : _lastScanTime.ToString("HH:mm:ss");
            string backupText = _lastBackupTime == DateTime.MinValue ? "暂无" : _lastBackupTime.ToString("HH:mm:ss") + " " + _lastBackupFileName;
            return "运行中    最近扫描：" + scanText +
                "    当前打开：" + _lastScanDocumentCount.ToString() + " 个" +
                "    今日备份：" + GetTodayBackupCount().ToString() + " 条" +
                "    最近备份：" + backupText +
                "    格式：" + BackupFormatDescription +
                "    位置：" + BackupLocationDescription +
                "    空间保护：" + SpaceLimitDescription;
        }

        public string BuildHealthReport()
        {
            StringBuilder builder = new StringBuilder();
            List<string> activeRoots = GetActiveBackupRoots();

            builder.AppendLine("泽PPT备份助手 一键体检");
            builder.AppendLine("======================");
            builder.AppendLine("体检时间：" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            builder.AppendLine("程序版本：v" + GetVersionText());
            builder.AppendLine("运行状态：运行中");
            builder.AppendLine("启用格式：" + BackupFormatDescription);
            builder.AppendLine("备份位置：" + BackupLocationDescription);
            builder.AppendLine("自动清理：" + AutoCleanDescription);
            builder.AppendLine("磁盘空间保护：" + SpaceLimitDescription);
            builder.AppendLine("开机自启：" + (IsAutoStartEnabled() ? "已开启" : "未开启"));
            builder.AppendLine("备份提示：" + (_successTipEnabled ? "已开启" : "未开启"));
            builder.AppendLine("免责声明：" + (_disclaimerAccepted ? "已确认" : "未确认"));
            builder.AppendLine();

            builder.AppendLine("当前生效备份位置：");
            if (activeRoots.Count == 0)
            {
                builder.AppendLine("  未找到可用备份位置");
            }
            else
            {
                for (int i = 0; i < activeRoots.Count; i++)
                {
                    builder.AppendLine("  " + activeRoots[i]);
                }
            }
            builder.AppendLine();

            builder.AppendLine("备份位置写入检查：");
            string[] roots = GetAllBackupRoots();
            for (int i = 0; i < roots.Length; i++)
            {
                builder.AppendLine("  " + TestBackupRootForReport(roots[i]));
            }
            builder.AppendLine();

            builder.AppendLine("日志与记录：");
            builder.AppendLine("  日志文件：" + LogFilePath + (File.Exists(LogFilePath) ? "（存在）" : "（未生成）"));
            builder.AppendLine("  记录文件：" + IndexFilePath + (File.Exists(IndexFilePath) ? "（存在）" : "（未生成）"));
            builder.AppendLine("  设置文件：" + FindSettingsPath());
            builder.AppendLine("  今日备份：" + GetTodayBackupCount().ToString() + " 条");
            List<BackupRecord> recent = GetRecentBackups(1);
            if (recent.Count > 0)
            {
                builder.AppendLine("  最近备份：" + recent[0].TimeText + "  " + recent[0].FileName + "  " + recent[0].BackupStatusText);
            }
            else
            {
                builder.AppendLine("  最近备份：暂无");
            }
            builder.AppendLine();

            builder.AppendLine("扫描状态：");
            builder.AppendLine("  最近扫描：" + (_lastScanTime == DateTime.MinValue ? "尚未扫描" : _lastScanTime.ToString("yyyy-MM-dd HH:mm:ss")));
            builder.AppendLine("  最近扫描发现文件：" + _lastScanDocumentCount.ToString() + " 个");
            builder.AppendLine("  最近成功备份：" + (_lastBackupTime == DateTime.MinValue ? "暂无" : _lastBackupTime.ToString("yyyy-MM-dd HH:mm:ss") + "  " + _lastBackupFileName));
            builder.AppendLine();

            builder.AppendLine("结论：");
            if (activeRoots.Count > 0 && (_backupPptEnabled || _backupWordEnabled || _backupPdfEnabled))
            {
                builder.AppendLine("  基础环境正常。打开支持格式文件后，程序会自动备份到当前生效备份位置。");
            }
            else
            {
                builder.AppendLine("  需要检查备份位置或备份格式设置。");
            }

            return builder.ToString();
        }

        private string TestBackupRootForReport(string root)
        {
            try
            {
                string driveRoot = Path.GetPathRoot(root);
                if (string.IsNullOrEmpty(driveRoot) || !Directory.Exists(driveRoot))
                {
                    return root + "：不可用，盘符不存在";
                }

                DriveInfo drive = new DriveInfo(driveRoot);
                if (!drive.IsReady)
                {
                    return root + "：不可用，磁盘未就绪";
                }

                Directory.CreateDirectory(root);
                string logDir = Path.Combine(root, "日志");
                Directory.CreateDirectory(logDir);
                string testPath = Path.Combine(logDir, ".write-test-" + Guid.NewGuid().ToString("N") + ".tmp");
                File.WriteAllText(testPath, "ok", Encoding.UTF8);
                File.Delete(testPath);

                long used = GetRootBackupBytes(root);
                return root + "：可写，备份占用 " + FormatSize(used) + "，磁盘剩余 " + FormatSize(drive.AvailableFreeSpace);
            }
            catch (Exception ex)
            {
                return root + "：不可用，" + ex.Message;
            }
        }

        private int GetTodayBackupCount()
        {
            int count = 0;
            try
            {
                if (!File.Exists(IndexFilePath))
                {
                    return 0;
                }

                string[] lines = File.ReadAllLines(IndexFilePath, Encoding.UTF8);
                DateTime today = DateTime.Today;
                for (int i = 0; i < lines.Length; i++)
                {
                    BackupRecord record = ParseBackupRecord(lines[i]);
                    if (record != null && record.Time.Date == today)
                    {
                        count++;
                    }
                }
            }
            catch
            {
            }
            return count;
        }

        private static long GetRootBackupBytes(string root)
        {
            long total = 0;
            try
            {
                List<BackupFileEntry> files = CollectBackupFiles(root);
                for (int i = 0; i < files.Count; i++)
                {
                    total += files[i].SizeBytes;
                }
            }
            catch
            {
            }
            return total;
        }

        private static string GetVersionText()
        {
            try
            {
                object[] attributes = Assembly.GetExecutingAssembly().GetCustomAttributes(typeof(AssemblyInformationalVersionAttribute), false);
                if (attributes.Length > 0)
                {
                    AssemblyInformationalVersionAttribute info = attributes[0] as AssemblyInformationalVersionAttribute;
                    if (info != null && !string.IsNullOrEmpty(info.InformationalVersion))
                    {
                        return info.InformationalVersion;
                    }
                }
            }
            catch
            {
            }
            return "5.0";
        }

        public int DeleteBackupRecord(BackupRecord record, bool deleteFiles)
        {
            if (record == null || string.IsNullOrEmpty(record.Fingerprint))
            {
                return 0;
            }

            _suppressedFingerprints.Add(record.Fingerprint);

            int deletedFiles = 0;
            if (deleteFiles)
            {
                string[] paths = SplitSavedPaths(record.SavedPaths);
                for (int i = 0; i < paths.Length; i++)
                {
                    try
                    {
                        if (!string.IsNullOrWhiteSpace(paths[i]) && File.Exists(paths[i]))
                        {
                            File.Delete(paths[i]);
                            deletedFiles++;
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("删除备份文件失败：" + paths[i] + "。原因：" + ex.Message);
                    }
                }
            }

            RemoveFingerprintsFromIndex(new string[] { record.Fingerprint });
            _backedUpFingerprints.Remove(record.Fingerprint);
            _fingerprintSavedPaths.Remove(record.Fingerprint);
            if (!string.IsNullOrEmpty(record.ContentHash))
            {
                _contentHashSavedPaths.Remove(record.ContentHash);
            }
            _recentNameSizeEntries.Remove(MakeNameSizeKey(record.Kind, record.FileName, record.SizeBytes));
            Log("删除备份记录：" + record.FileName + "，删除文件数：" + deletedFiles.ToString() + "。本次运行内已临时忽略同一版本源文件。");
            return deletedFiles;
        }

        public int CleanMissingBackupRecords()
        {
            int removed = 0;

            try
            {
                if (!File.Exists(IndexFilePath))
                {
                    return 0;
                }

                string[] lines = File.ReadAllLines(IndexFilePath, Encoding.UTF8);
                List<string> kept = new List<string>();
                for (int i = 0; i < lines.Length; i++)
                {
                    BackupRecord record = ParseBackupRecord(lines[i]);
                    if (record == null)
                    {
                        continue;
                    }

                    if (record.HasBackupFiles)
                    {
                        kept.Add(lines[i]);
                    }
                    else
                    {
                        removed++;
                    }
                }

                File.WriteAllLines(IndexFilePath, kept.ToArray(), Encoding.UTF8);
                LoadIndex();
                Log("清理失效记录完成，移除：" + removed.ToString() + " 条。");
            }
            catch (Exception ex)
            {
                Log("清理失效记录失败：" + ex.Message);
            }

            return removed;
        }

        public void ClearBackupRecords()
        {
            try
            {
                EnsureLogDirectory();
                File.WriteAllText(IndexFilePath, "", Encoding.UTF8);
                _backedUpFingerprints.Clear();
                _fingerprintSavedPaths.Clear();
                _contentHashSavedPaths.Clear();
                _recentNameSizeEntries.Clear();
                _suppressedFingerprints.Clear();
                Log("备份记录已清空。");
            }
            catch (Exception ex)
            {
                Log("清空备份记录失败：" + ex.Message);
            }
        }

        public void ClearLogFile()
        {
            try
            {
                EnsureLogDirectory();
                File.WriteAllText(LogFilePath, "", Encoding.UTF8);
            }
            catch
            {
            }
        }

        private void HandleDocument(OpenDocument opened, DateTime nowUtc)
        {
            if (opened == null || string.IsNullOrEmpty(opened.Path))
            {
                return;
            }

            string path;
            try
            {
                path = Path.GetFullPath(opened.Path);
            }
            catch
            {
                return;
            }

            if (!IsAllowedDocument(path) || !IsKindEnabled(opened.Kind) || IsInsideBackupRoot(path) || !File.Exists(path))
            {
                return;
            }

            FileInfo info;
            try
            {
                info = new FileInfo(path);
            }
            catch
            {
                return;
            }

            PendingDocument pending;
            if (!_pending.TryGetValue(path, out pending))
            {
                pending = new PendingDocument();
                pending.Path = path;
                pending.LastLength = info.Length;
                pending.LastWriteUtcTicks = info.LastWriteTimeUtc.Ticks;
                pending.StableSinceUtc = nowUtc;
                _pending[path] = pending;
                return;
            }

            if (pending.LastLength != info.Length || pending.LastWriteUtcTicks != info.LastWriteTimeUtc.Ticks)
            {
                pending.LastLength = info.Length;
                pending.LastWriteUtcTicks = info.LastWriteTimeUtc.Ticks;
                pending.StableSinceUtc = nowUtc;
                return;
            }

            if ((nowUtc - pending.StableSinceUtc).TotalSeconds < 6)
            {
                return;
            }

            string fingerprint = MakeFingerprint(path, info.Length, info.LastWriteTimeUtc.Ticks);
            if (_suppressedFingerprints.Contains(fingerprint))
            {
                return;
            }

            if (IsFingerprintStillBackedUp(fingerprint))
            {
                return;
            }

            string contentHash = ComputeContentHash(path);
            if (!string.IsNullOrEmpty(contentHash) && IsContentHashStillBackedUp(contentHash))
            {
                Log("内容相同，跳过重复备份：" + path);
                return;
            }

            if (IsRecentNameSizeStillBackedUp(opened.Kind, Path.GetFileName(path), info.Length))
            {
                Log("短时间内同名同大小，跳过重复备份：" + path);
                return;
            }

            BackupFile(path, info, opened.AppName, opened.Kind, fingerprint, contentHash);
        }

        private List<OpenDocument> GetOpenPresentations()
        {
            List<OpenDocument> result = new List<OpenDocument>();
            Dictionary<string, bool> seen = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < PresentationProgIds.Length; i++)
            {
                string progId = PresentationProgIds[i];
                object rawApp = null;

                try
                {
                    rawApp = Marshal.GetActiveObject(progId);
                }
                catch
                {
                    continue;
                }

                try
                {
                    dynamic app = rawApp;
                    dynamic presentations = app.Presentations;
                    int count = Convert.ToInt32(presentations.Count);

                    for (int index = 1; index <= count; index++)
                    {
                        try
                        {
                            dynamic item = presentations.Item(index);
                            string fullName = Convert.ToString(item.FullName);
                            AddOpenDocument(result, seen, fullName, "PPT", FriendlyAppName(progId));
                        }
                        catch
                        {
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogProgIdErrorOnce(progId, ex);
                }
                finally
                {
                    ReleaseCom(rawApp);
                }
            }

            return result;
        }

        private List<OpenDocument> GetOpenWordDocuments()
        {
            List<OpenDocument> result = new List<OpenDocument>();
            Dictionary<string, bool> seen = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < WordProgIds.Length; i++)
            {
                string progId = WordProgIds[i];
                object rawApp = null;

                try
                {
                    rawApp = Marshal.GetActiveObject(progId);
                }
                catch
                {
                    continue;
                }

                try
                {
                    dynamic app = rawApp;
                    dynamic documents = app.Documents;
                    int count = Convert.ToInt32(documents.Count);

                    for (int index = 1; index <= count; index++)
                    {
                        try
                        {
                            dynamic item = documents.Item(index);
                            string fullName = Convert.ToString(item.FullName);
                            AddOpenDocument(result, seen, fullName, "Word", FriendlyWordAppName(progId));
                        }
                        catch
                        {
                        }
                    }
                }
                catch (Exception ex)
                {
                    LogProgIdErrorOnce(progId, ex);
                }
                finally
                {
                    ReleaseCom(rawApp);
                }
            }

            return result;
        }

        private List<OpenDocument> GetOpenPdfDocuments()
        {
            List<OpenDocument> result = new List<OpenDocument>();
            Dictionary<string, bool> seen = new Dictionary<string, bool>(StringComparer.OrdinalIgnoreCase);

            try
            {
                using (ManagementObjectSearcher searcher = new ManagementObjectSearcher("SELECT Name, CommandLine FROM Win32_Process"))
                {
                    foreach (ManagementObject process in searcher.Get())
                    {
                        string name = Convert.ToString(process["Name"]);
                        string commandLine = Convert.ToString(process["CommandLine"]);
                        if (string.IsNullOrEmpty(commandLine) || !LooksLikePdfReader(name, commandLine))
                        {
                            continue;
                        }

                        List<string> paths = ExtractPdfPaths(commandLine);
                        for (int i = 0; i < paths.Count; i++)
                        {
                            AddOpenDocument(result, seen, paths[i], "PDF", "PDF");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("读取 PDF 进程信息失败：" + ex.Message);
            }

            AddPdfDocumentsFromBrowserTitles(result, seen);
            return result;
        }

        private void AddPdfDocumentsFromBrowserTitles(List<OpenDocument> result, Dictionary<string, bool> seen)
        {
            try
            {
                Process[] browsers = Process.GetProcessesByName("QQBrowser");
                for (int i = 0; i < browsers.Length; i++)
                {
                    string title = "";
                    try
                    {
                        title = browsers[i].MainWindowTitle;
                    }
                    catch
                    {
                    }

                    string candidateTitle = PdfTitleCandidateFromWindowTitle(title);
                    if (string.IsNullOrEmpty(candidateTitle))
                    {
                        continue;
                    }

                    string path = ResolvePdfByTitle(candidateTitle);
                    if (!string.IsNullOrEmpty(path))
                    {
                        AddOpenDocument(result, seen, path, "PDF", "QQ浏览器");
                    }
                }
            }
            catch (Exception ex)
            {
                Log("读取 QQ 浏览器窗口标题失败：" + ex.Message);
            }
        }

        private static string PdfTitleCandidateFromWindowTitle(string title)
        {
            if (string.IsNullOrWhiteSpace(title))
            {
                return "";
            }

            string cleaned = title.Trim();
            string[] suffixes = new string[]
            {
                " - QQ浏览器",
                "- QQ浏览器",
                " - QQBrowser",
                "- QQBrowser",
                " - PDF阅读器",
                "- PDF阅读器"
            };

            for (int i = 0; i < suffixes.Length; i++)
            {
                if (cleaned.EndsWith(suffixes[i], StringComparison.OrdinalIgnoreCase))
                {
                    cleaned = cleaned.Substring(0, cleaned.Length - suffixes[i].Length).Trim();
                    break;
                }
            }

            if (cleaned.Length == 0 || cleaned.Length > 180)
            {
                return "";
            }

            string lower = cleaned.ToLowerInvariant();
            if (lower == "qq浏览器" || lower == "新标签页" || lower == "new tab" || lower.StartsWith("http://") || lower.StartsWith("https://"))
            {
                return "";
            }

            return cleaned;
        }

        private string ResolvePdfByTitle(string title)
        {
            TitleSearchCacheEntry cache;
            if (_pdfTitleCache.TryGetValue(title, out cache) && (DateTime.UtcNow - cache.CheckedUtc).TotalSeconds < 45)
            {
                return cache.Path;
            }

            string path = SearchPdfByTitle(title);
            cache = new TitleSearchCacheEntry();
            cache.CheckedUtc = DateTime.UtcNow;
            cache.Path = path;
            _pdfTitleCache[title] = cache;
            return path;
        }

        private string SearchPdfByTitle(string title)
        {
            string exactName = title.EndsWith(".pdf", StringComparison.OrdinalIgnoreCase) ? title : title + ".pdf";
            List<string> roots = BuildPdfSearchRoots();

            for (int i = 0; i < roots.Count; i++)
            {
                string found = SearchDirectoryForPdf(roots[i], title, exactName);
                if (!string.IsNullOrEmpty(found))
                {
                    Log("通过 QQ 浏览器窗口标题找到 PDF：" + title + " -> " + found);
                    return found;
                }
            }

            return "";
        }

        private static List<string> BuildPdfSearchRoots()
        {
            List<string> roots = new List<string>();
            string userProfile = Environment.GetEnvironmentVariable("USERPROFILE");
            AddRoot(roots, Path.Combine(userProfile ?? "", "Desktop"));
            AddRoot(roots, Path.Combine(userProfile ?? "", "Downloads"));
            AddRoot(roots, Path.Combine(userProfile ?? "", "Documents"));

            DriveInfo[] drives = DriveInfo.GetDrives();
            for (int i = 0; i < drives.Length; i++)
            {
                try
                {
                    if (!drives[i].IsReady)
                    {
                        continue;
                    }

                    DriveType type = drives[i].DriveType;
                    if (type == DriveType.Removable || type == DriveType.Fixed)
                    {
                        AddRoot(roots, drives[i].RootDirectory.FullName);
                    }
                }
                catch
                {
                }
            }

            return roots;
        }

        private static void AddRoot(List<string> roots, string path)
        {
            if (string.IsNullOrEmpty(path) || !Directory.Exists(path))
            {
                return;
            }

            string full;
            try
            {
                full = Path.GetFullPath(path);
            }
            catch
            {
                return;
            }

            for (int i = 0; i < roots.Count; i++)
            {
                if (string.Equals(roots[i], full, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }
            roots.Add(full);
        }

        private static string SearchDirectoryForPdf(string root, string title, string exactName)
        {
            Queue<SearchDir> queue = new Queue<SearchDir>();
            queue.Enqueue(new SearchDir(root, 0));
            int visited = 0;

            while (queue.Count > 0 && visited < 2500)
            {
                SearchDir current = queue.Dequeue();
                visited++;

                try
                {
                    string exactPath = Path.Combine(current.Path, exactName);
                    if (File.Exists(exactPath))
                    {
                        return exactPath;
                    }

                    string[] pdfs = Directory.GetFiles(current.Path, "*.pdf");
                    for (int i = 0; i < pdfs.Length; i++)
                    {
                        string name = Path.GetFileNameWithoutExtension(pdfs[i]);
                        if (string.Equals(name, title, StringComparison.OrdinalIgnoreCase))
                        {
                            return pdfs[i];
                        }
                    }

                    if (current.Depth >= 5)
                    {
                        continue;
                    }

                    string[] dirs = Directory.GetDirectories(current.Path);
                    for (int i = 0; i < dirs.Length; i++)
                    {
                        if (!ShouldSkipDirectory(dirs[i]))
                        {
                            queue.Enqueue(new SearchDir(dirs[i], current.Depth + 1));
                        }
                    }
                }
                catch
                {
                }
            }

            return "";
        }

        private static bool ShouldSkipDirectory(string path)
        {
            string name = Path.GetFileName(path).ToLowerInvariant();
            string full = path.ToLowerInvariant();

            if (name == "windows" || name == "program files" || name == "program files (x86)" || name == "node_modules" || name == ".git")
            {
                return true;
            }

            if (full.IndexOf("\\appdata\\", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            if (full.IndexOf("\\泽宁pppppppptttt备份", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return true;
            }

            return false;
        }

        private static bool LooksLikePdfReader(string name, string commandLine)
        {
            string combined = (name + " " + commandLine).ToLowerInvariant();
            if (combined.IndexOf(".pdf", StringComparison.OrdinalIgnoreCase) < 0)
            {
                return false;
            }

            return combined.IndexOf("acrobat") >= 0
                || combined.IndexOf("acrord") >= 0
                || combined.IndexOf("foxit") >= 0
                || combined.IndexOf("wps") >= 0
                || combined.IndexOf("kwps") >= 0
                || combined.IndexOf("edge") >= 0
                || combined.IndexOf("chrome") >= 0
                || combined.IndexOf("browser") >= 0
                || combined.IndexOf("pdf") >= 0;
        }

        private static List<string> ExtractPdfPaths(string commandLine)
        {
            List<string> result = new List<string>();
            AddPdfMatches(result, commandLine);

            MatchCollection fileUris = Regex.Matches(commandLine, @"file:///[^\s""]+?\.pdf", RegexOptions.IgnoreCase);
            for (int i = 0; i < fileUris.Count; i++)
            {
                try
                {
                    string uriText = fileUris[i].Value;
                    Uri uri = new Uri(uriText);
                    AddPdfMatches(result, Uri.UnescapeDataString(uri.LocalPath));
                }
                catch
                {
                }
            }

            return result;
        }

        private static void AddPdfMatches(List<string> result, string text)
        {
            if (string.IsNullOrEmpty(text))
            {
                return;
            }

            MatchCollection quoted = Regex.Matches(text, "\"([^\"]+?\\.pdf)\"", RegexOptions.IgnoreCase);
            for (int i = 0; i < quoted.Count; i++)
            {
                AddIfPdfPath(result, quoted[i].Groups[1].Value);
            }

            MatchCollection plain = Regex.Matches(text, @"(?<![A-Za-z0-9])([A-Za-z]:\\[^\r\n\t""<>|]+?\.pdf)", RegexOptions.IgnoreCase);
            for (int i = 0; i < plain.Count; i++)
            {
                AddIfPdfPath(result, plain[i].Groups[1].Value.Trim());
            }
        }

        private static void AddIfPdfPath(List<string> result, string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return;
            }

            string path = value.Trim().Trim('"');
            try
            {
                path = Uri.UnescapeDataString(path);
            }
            catch
            {
            }

            try
            {
                path = Path.GetFullPath(path);
            }
            catch
            {
                return;
            }

            if (string.Equals(Path.GetExtension(path), ".pdf", StringComparison.OrdinalIgnoreCase) && File.Exists(path))
            {
                for (int i = 0; i < result.Count; i++)
                {
                    if (string.Equals(result[i], path, StringComparison.OrdinalIgnoreCase))
                    {
                        return;
                    }
                }
                result.Add(path);
            }
        }

        private static void AddOpenDocument(List<OpenDocument> result, Dictionary<string, bool> seen, string fullName, string kind, string appName)
        {
            if (string.IsNullOrEmpty(fullName))
            {
                return;
            }

            string path;
            try
            {
                path = Path.GetFullPath(fullName);
            }
            catch
            {
                return;
            }

            if (!IsAllowedDocument(path) || IsInsideBackupRoot(path) || !File.Exists(path) || seen.ContainsKey(path))
            {
                return;
            }

            seen[path] = true;
            result.Add(new OpenDocument(path, kind, appName));
        }

        private static string FriendlyAppName(string progId)
        {
            if (progId.IndexOf("PowerPoint", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "PowerPoint";
            }
            return "WPS";
        }

        private static string FriendlyWordAppName(string progId)
        {
            if (progId.IndexOf("Word", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                return "Word";
            }
            return "WPS";
        }

        private void BackupFile(string sourcePath, FileInfo sourceInfo, string appName, string kind, string fingerprint, string contentHash)
        {
            List<string> saved = new List<string>();
            List<string> failed = new List<string>();
            string fileName = Path.GetFileName(sourcePath);
            string dateFolder = DateTime.Now.ToString("yyyy-MM-dd");

            List<string> backupRoots = GetActiveBackupRoots();
            for (int i = 0; i < backupRoots.Count; i++)
            {
                string root = backupRoots[i];
                try
                {
                    string targetDir = Path.Combine(root, dateFolder);
                    Directory.CreateDirectory(targetDir);
                    string targetPath = NextAvailablePath(targetDir, fileName);
                    CopyShared(sourcePath, targetPath);
                    File.SetLastWriteTimeUtc(targetPath, sourceInfo.LastWriteTimeUtc);
                    saved.Add(targetPath);
                }
                catch (Exception ex)
                {
                    failed.Add(root + "：" + ex.Message);
                }
            }

            if (backupRoots.Count == 0)
            {
                failed.Add("D/E/C 备份目录都不可用");
            }

            if (saved.Count > 0)
            {
                _suppressedFingerprints.Remove(fingerprint);
                _backedUpFingerprints.Add(fingerprint);
                _fingerprintSavedPaths[fingerprint] = string.Join(" ; ", saved.ToArray());
                if (!string.IsNullOrEmpty(contentHash))
                {
                    _contentHashSavedPaths[contentHash] = string.Join(" ; ", saved.ToArray());
                }
                AddRecentNameSizeEntry(kind, Path.GetFileName(sourcePath), sourceInfo.Length, DateTime.Now, string.Join(" ; ", saved.ToArray()), fingerprint);
                AppendIndex(fingerprint, sourcePath, sourceInfo, kind, contentHash, saved);
                _lastBackupTime = DateTime.Now;
                _lastBackupFileName = Path.GetFileName(sourcePath);
                Log("备份成功 [" + appName + "/" + kind + "] " + sourcePath + " -> " + string.Join(" ; ", saved.ToArray()));
                RaiseBackupSucceeded(Path.GetFileName(sourcePath), string.Join(" ; ", saved.ToArray()));
                CleanupResult cleanup = RunSpaceLimitCleanNow();
                if (cleanup.DeletedFiles > 0)
                {
                    Log("磁盘空间保护已清理文件：" + cleanup.DeletedFiles.ToString() + " 个，释放：" + FormatSize(cleanup.FreedBytes) + "。");
                }
            }
            else
            {
                Log("备份失败 [" + appName + "/" + kind + "] " + sourcePath + "。原因：" + string.Join(" ; ", failed.ToArray()));
            }
        }

        private static void CopyShared(string sourcePath, string targetPath)
        {
            using (FileStream source = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            using (FileStream target = new FileStream(targetPath, FileMode.CreateNew, FileAccess.Write, FileShare.None))
            {
                byte[] buffer = new byte[1024 * 1024];
                while (true)
                {
                    int read = source.Read(buffer, 0, buffer.Length);
                    if (read <= 0)
                    {
                        break;
                    }
                    target.Write(buffer, 0, read);
                }
            }
        }

        private static string NextAvailablePath(string directory, string fileName)
        {
            string name = Path.GetFileNameWithoutExtension(fileName);
            string extension = Path.GetExtension(fileName);
            string candidate = Path.Combine(directory, fileName);
            int index = 2;

            while (File.Exists(candidate))
            {
                candidate = Path.Combine(directory, name + "_" + index.ToString() + extension);
                index++;
            }

            return candidate;
        }

        private void EnsureBackupRoots()
        {
            GetActiveBackupRoots();
        }

        private List<string> GetActiveBackupRoots()
        {
            List<string> roots = new List<string>();

            string[] primaryRoots = GetPrimaryBackupRoots();
            for (int i = 0; i < primaryRoots.Length; i++)
            {
                if (EnsureBackupRoot(primaryRoots[i]))
                {
                    AddUniqueRoot(roots, primaryRoots[i]);
                }
            }

            if (roots.Count == 0 && !string.IsNullOrWhiteSpace(_customBackupRoot))
            {
                for (int i = 0; i < BackupRoots.Length; i++)
                {
                    if (EnsureBackupRoot(BackupRoots[i]))
                    {
                        AddUniqueRoot(roots, BackupRoots[i]);
                    }
                }
            }

            if (roots.Count == 0)
            {
                string[] fallbackRoots = GetFallbackBackupRoots();
                for (int i = 0; i < fallbackRoots.Length; i++)
                {
                    if (EnsureBackupRoot(fallbackRoots[i]))
                    {
                        AddUniqueRoot(roots, fallbackRoots[i]);
                        string key = fallbackRoots[i] + "|fallback";
                        if (!_rootStatusLogged.Contains(key))
                        {
                            _rootStatusLogged.Add(key);
                            Log("D/E 备份盘都不可用，已启用 C 盘备用备份目录：" + fallbackRoots[i]);
                        }
                        break;
                    }
                }
            }

            return roots;
        }

        private string[] GetPrimaryBackupRoots()
        {
            if (!string.IsNullOrWhiteSpace(_customBackupRoot))
            {
                return new string[] { _customBackupRoot };
            }
            return BackupRoots;
        }

        private bool EnsureBackupRoot(string root)
        {
            try
            {
                string driveRoot = Path.GetPathRoot(root);
                if (string.IsNullOrEmpty(driveRoot) || !Directory.Exists(driveRoot))
                {
                    LogRootOnce(root, "备份盘不存在，暂时跳过：" + root);
                    return false;
                }

                DriveInfo drive = new DriveInfo(driveRoot);
                if (!drive.IsReady)
                {
                    LogRootOnce(root, "备份盘未就绪，暂时跳过：" + root);
                    return false;
                }

                Directory.CreateDirectory(root);
                Directory.CreateDirectory(Path.Combine(root, DateTime.Now.ToString("yyyy-MM-dd")));

                if (!_rootStatusLogged.Contains(root + "|ok"))
                {
                    _rootStatusLogged.Add(root + "|ok");
                    Log("备份目录已就绪：" + root);
                }
                return true;
            }
            catch (Exception ex)
            {
                LogRootOnce(root, "备份目录不可用：" + root + "。原因：" + ex.Message);
                return false;
            }
        }

        private void LogRootOnce(string root, string message)
        {
            string key = root + "|bad|" + message;
            if (!_rootStatusLogged.Contains(key))
            {
                _rootStatusLogged.Add(key);
                Log(message);
            }
        }

        private void LogProgIdErrorOnce(string progId, Exception ex)
        {
            if (!_progIdErrorLogged.Contains(progId))
            {
                _progIdErrorLogged.Add(progId);
                Log("读取 " + progId + " 打开的文件失败：" + ex.Message);
            }
        }

        private string GetDataRoot()
        {
            if (!string.IsNullOrWhiteSpace(_customBackupRoot) && CanUseBackupRoot(_customBackupRoot))
            {
                _dataRoot = _customBackupRoot;
                return _dataRoot;
            }

            if (!string.IsNullOrEmpty(_dataRoot) && CanUseBackupRoot(_dataRoot))
            {
                return _dataRoot;
            }

            string[] roots = GetAllBackupRoots();
            for (int i = 0; i < roots.Length; i++)
            {
                if (CanUseBackupRoot(roots[i]))
                {
                    _dataRoot = roots[i];
                    return _dataRoot;
                }
            }

            _dataRoot = BackupRoots[0];
            return _dataRoot;
        }

        private static bool CanUseBackupRoot(string root)
        {
            try
            {
                string driveRoot = Path.GetPathRoot(root);
                if (string.IsNullOrEmpty(driveRoot) || !Directory.Exists(driveRoot))
                {
                    return false;
                }

                DriveInfo drive = new DriveInfo(driveRoot);
                if (!drive.IsReady)
                {
                    return false;
                }

                Directory.CreateDirectory(root);
                Directory.CreateDirectory(Path.Combine(root, "日志"));
                return true;
            }
            catch
            {
                return false;
            }
        }

        private static string[] GetAllBackupRoots()
        {
            List<string> roots = new List<string>();
            AddUniqueRoot(roots, s_customBackupRoot);
            for (int i = 0; i < BackupRoots.Length; i++)
            {
                AddUniqueRoot(roots, BackupRoots[i]);
            }

            string[] fallbackRoots = GetFallbackBackupRoots();
            for (int i = 0; i < fallbackRoots.Length; i++)
            {
                AddUniqueRoot(roots, fallbackRoots[i]);
            }

            return roots.ToArray();
        }

        private static string[] GetFallbackBackupRoots()
        {
            List<string> roots = new List<string>();
            AddUniqueRoot(roots, Path.Combine(GetCDriveRoot(), BackupFolderName));

            string documents = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            if (string.IsNullOrEmpty(documents))
            {
                string userProfile = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile);
                if (!string.IsNullOrEmpty(userProfile))
                {
                    documents = Path.Combine(userProfile, "Documents");
                }
            }
            if (!string.IsNullOrEmpty(documents))
            {
                AddUniqueRoot(roots, Path.Combine(documents, BackupFolderName));
            }

            return roots.ToArray();
        }

        private static string GetCDriveRoot()
        {
            if (Directory.Exists(@"C:\"))
            {
                return @"C:\";
            }

            string systemRoot = Path.GetPathRoot(Environment.SystemDirectory);
            if (!string.IsNullOrEmpty(systemRoot))
            {
                return systemRoot;
            }

            return @"C:\";
        }

        private static void AddUniqueRoot(List<string> roots, string root)
        {
            if (string.IsNullOrWhiteSpace(root))
            {
                return;
            }

            string full;
            try
            {
                full = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
            }
            catch
            {
                return;
            }

            for (int i = 0; i < roots.Count; i++)
            {
                if (string.Equals(roots[i].TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar), full, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }

            roots.Add(full);
        }

        private static void ReleaseCom(object rawApp)
        {
            if (rawApp != null && Marshal.IsComObject(rawApp))
            {
                try
                {
                    Marshal.ReleaseComObject(rawApp);
                }
                catch
                {
                }
            }
        }

        private void EnsureLogDirectory()
        {
            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(LogFilePath));
            }
            catch
            {
            }
        }

        private void LoadSettings()
        {
            _successTipEnabled = false;
            _backupPptEnabled = true;
            _backupWordEnabled = true;
            _backupPdfEnabled = true;
            _disclaimerAccepted = false;
            _customBackupRoot = "";
            s_customBackupRoot = "";
            _autoCleanDays = 0;
            _maxBackupSizeMb = DefaultMaxBackupSizeMb;
            _lastAutoCleanDate = DateTime.MinValue;

            try
            {
                string settingsPath = FindSettingsPath();
                if (string.IsNullOrEmpty(settingsPath))
                {
                    SaveSettings();
                    return;
                }

                string[] lines = File.ReadAllLines(settingsPath, Encoding.UTF8);
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i].Trim();
                    if (line.Length == 0 || line.StartsWith("#"))
                    {
                        continue;
                    }

                    int equals = line.IndexOf('=');
                    if (equals <= 0)
                    {
                        continue;
                    }

                    string key = line.Substring(0, equals).Trim();
                    string value = line.Substring(equals + 1).Trim();
                    if (string.Equals(key, "SuccessTipEnabled", StringComparison.OrdinalIgnoreCase))
                    {
                        bool parsed;
                        if (bool.TryParse(value, out parsed))
                        {
                            _successTipEnabled = parsed;
                        }
                    }
                    else if (string.Equals(key, "BackupPptEnabled", StringComparison.OrdinalIgnoreCase))
                    {
                        bool parsed;
                        if (bool.TryParse(value, out parsed))
                        {
                            _backupPptEnabled = parsed;
                        }
                    }
                    else if (string.Equals(key, "BackupWordEnabled", StringComparison.OrdinalIgnoreCase))
                    {
                        bool parsed;
                        if (bool.TryParse(value, out parsed))
                        {
                            _backupWordEnabled = parsed;
                        }
                    }
                    else if (string.Equals(key, "BackupPdfEnabled", StringComparison.OrdinalIgnoreCase))
                    {
                        bool parsed;
                        if (bool.TryParse(value, out parsed))
                        {
                            _backupPdfEnabled = parsed;
                        }
                    }
                    else if (string.Equals(key, "DisclaimerAccepted", StringComparison.OrdinalIgnoreCase))
                    {
                        bool parsed;
                        if (bool.TryParse(value, out parsed))
                        {
                            _disclaimerAccepted = parsed;
                        }
                    }
                    else if (string.Equals(key, "CustomBackupRoot", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            _customBackupRoot = string.IsNullOrWhiteSpace(value) ? "" : Path.GetFullPath(value);
                            s_customBackupRoot = _customBackupRoot;
                            _dataRoot = null;
                        }
                        catch
                        {
                            _customBackupRoot = "";
                            s_customBackupRoot = "";
                        }
                    }
                    else if (string.Equals(key, "AutoCleanDays", StringComparison.OrdinalIgnoreCase))
                    {
                        int parsedDays;
                        if (int.TryParse(value, out parsedDays) && (parsedDays == 0 || parsedDays == 30 || parsedDays == 60 || parsedDays == 90))
                        {
                            _autoCleanDays = parsedDays;
                        }
                    }
                    else if (string.Equals(key, "MaxBackupSizeMb", StringComparison.OrdinalIgnoreCase))
                    {
                        int parsedMegabytes;
                        if (int.TryParse(value, out parsedMegabytes) && (parsedMegabytes == 0 || parsedMegabytes == 2048 || parsedMegabytes == 5120 || parsedMegabytes == 10240 || parsedMegabytes == 20480 || parsedMegabytes == 51200))
                        {
                            _maxBackupSizeMb = parsedMegabytes;
                        }
                    }
                    else if (string.Equals(key, "LastAutoCleanDate", StringComparison.OrdinalIgnoreCase))
                    {
                        DateTime parsedDate;
                        if (DateTime.TryParse(value, out parsedDate))
                        {
                            _lastAutoCleanDate = parsedDate.Date;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log("读取设置失败：" + ex.Message);
            }

            if (!_backupPptEnabled && !_backupWordEnabled && !_backupPdfEnabled)
            {
                _backupPptEnabled = true;
            }
        }

        private void SaveSettings()
        {
            try
            {
                string text =
                    "SuccessTipEnabled=" + _successTipEnabled.ToString() + "\r\n" +
                    "BackupPptEnabled=" + _backupPptEnabled.ToString() + "\r\n" +
                    "BackupWordEnabled=" + _backupWordEnabled.ToString() + "\r\n" +
                    "BackupPdfEnabled=" + _backupPdfEnabled.ToString() + "\r\n" +
                    "DisclaimerAccepted=" + _disclaimerAccepted.ToString() + "\r\n" +
                    "CustomBackupRoot=" + (_customBackupRoot ?? "").Replace("\r", "").Replace("\n", "") + "\r\n" +
                    "AutoCleanDays=" + _autoCleanDays.ToString() + "\r\n" +
                    "MaxBackupSizeMb=" + _maxBackupSizeMb.ToString() + "\r\n" +
                    "LastAutoCleanDate=" + (_lastAutoCleanDate == DateTime.MinValue ? "" : _lastAutoCleanDate.ToString("yyyy-MM-dd")) + "\r\n";
                string[] paths = GetSettingsWritePaths();
                for (int i = 0; i < paths.Length; i++)
                {
                    try
                    {
                        string directory = Path.GetDirectoryName(paths[i]);
                        if (!string.IsNullOrEmpty(directory))
                        {
                            Directory.CreateDirectory(directory);
                        }
                        File.WriteAllText(paths[i], text, Encoding.UTF8);
                    }
                    catch
                    {
                    }
                }
            }
            catch (Exception ex)
            {
                Log("保存设置失败：" + ex.Message);
            }
        }

        private string FindSettingsPath()
        {
            string[] paths = GetSettingsCandidatePaths();
            for (int i = 0; i < paths.Length; i++)
            {
                if (File.Exists(paths[i]))
                {
                    return paths[i];
                }
            }
            return "";
        }

        private string[] GetSettingsCandidatePaths()
        {
            List<string> paths = new List<string>();
            AddUniqueRoot(paths, GetLocalSettingsPath());

            try
            {
                string exeDir = AppDomain.CurrentDomain.BaseDirectory;
                if (!string.IsNullOrEmpty(exeDir))
                {
                    AddUniqueRoot(paths, Path.Combine(exeDir, SettingsFileName));
                }
            }
            catch
            {
            }

            try
            {
                AddUniqueRoot(paths, SettingsFilePath);
            }
            catch
            {
            }

            string[] roots = GetAllBackupRoots();
            for (int i = 0; i < roots.Length; i++)
            {
                AddUniqueRoot(paths, Path.Combine(roots[i], "日志", SettingsFileName));
            }

            return paths.ToArray();
        }

        private string[] GetSettingsWritePaths()
        {
            List<string> paths = new List<string>();
            AddUniqueRoot(paths, GetLocalSettingsPath());
            AddUniqueRoot(paths, SettingsFilePath);
            return paths.ToArray();
        }

        private static string GetLocalSettingsPath()
        {
            return Path.Combine(AppDomain.CurrentDomain.BaseDirectory, SettingsFileName);
        }

        private int CleanOldBackupsIfNeeded(bool force)
        {
            if (_autoCleanDays <= 0)
            {
                return 0;
            }

            DateTime today = DateTime.Today;
            if (!force && _lastAutoCleanDate == today)
            {
                return 0;
            }

            int removedFolders = 0;
            DateTime keepFrom = today.AddDays(-(_autoCleanDays - 1));
            List<string> roots = GetActiveBackupRoots();
            for (int i = 0; i < roots.Count; i++)
            {
                string root = roots[i];
                try
                {
                    if (!Directory.Exists(root))
                    {
                        continue;
                    }

                    string[] directories = Directory.GetDirectories(root);
                    for (int j = 0; j < directories.Length; j++)
                    {
                        string directory = directories[j];
                        string name = Path.GetFileName(directory);
                        DateTime folderDate;
                        if (!DateTime.TryParseExact(name, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out folderDate))
                        {
                            continue;
                        }

                        if (folderDate >= keepFrom)
                        {
                            continue;
                        }

                        if (!IsPathUnderRoot(directory, root))
                        {
                            continue;
                        }

                        Directory.Delete(directory, true);
                        removedFolders++;
                    }
                }
                catch (Exception ex)
                {
                    Log("自动清理旧备份失败：" + root + "。原因：" + ex.Message);
                }
            }

            _lastAutoCleanDate = today;
            SaveSettings();

            if (force || removedFolders > 0)
            {
                Log("自动清理旧备份完成。保留最近 " + _autoCleanDays.ToString() + " 天，删除旧日期文件夹：" + removedFolders.ToString() + " 个。");
            }

            return removedFolders;
        }

        private CleanupResult CleanBackupsBySizeLimit()
        {
            CleanupResult result = new CleanupResult();
            if (_maxBackupSizeMb <= 0)
            {
                return result;
            }

            long limitBytes = _maxBackupSizeMb * 1024L * 1024L;
            string[] roots = GetAllBackupRoots();
            for (int i = 0; i < roots.Length; i++)
            {
                string root = roots[i];
                try
                {
                    if (!Directory.Exists(root))
                    {
                        continue;
                    }

                    List<BackupFileEntry> files = CollectBackupFiles(root);
                    long totalBytes = 0;
                    for (int j = 0; j < files.Count; j++)
                    {
                        totalBytes += files[j].SizeBytes;
                    }

                    if (totalBytes <= limitBytes)
                    {
                        continue;
                    }

                    files.Sort(delegate(BackupFileEntry left, BackupFileEntry right)
                    {
                        int compare = left.LastWriteUtc.CompareTo(right.LastWriteUtc);
                        if (compare != 0)
                        {
                            return compare;
                        }
                        return string.Compare(left.Path, right.Path, StringComparison.OrdinalIgnoreCase);
                    });

                    for (int j = 0; j < files.Count && totalBytes > limitBytes; j++)
                    {
                        BackupFileEntry file = files[j];
                        if (!IsPathUnderRoot(file.Path, root))
                        {
                            continue;
                        }

                        try
                        {
                            if (File.Exists(file.Path))
                            {
                                File.Delete(file.Path);
                                totalBytes -= file.SizeBytes;
                                result.DeletedFiles++;
                                result.FreedBytes += file.SizeBytes;
                            }
                        }
                        catch (Exception ex)
                        {
                            Log("磁盘空间保护删除文件失败：" + file.Path + "。原因：" + ex.Message);
                        }
                    }

                    result.RemovedFolders += RemoveEmptyDateFolders(root);
                }
                catch (Exception ex)
                {
                    Log("磁盘空间保护清理失败：" + root + "。原因：" + ex.Message);
                }
            }

            return result;
        }

        private static List<BackupFileEntry> CollectBackupFiles(string root)
        {
            List<BackupFileEntry> files = new List<BackupFileEntry>();
            string[] directories = Directory.GetDirectories(root);
            for (int i = 0; i < directories.Length; i++)
            {
                string name = Path.GetFileName(directories[i]);
                DateTime folderDate;
                if (!DateTime.TryParseExact(name, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out folderDate))
                {
                    continue;
                }

                string[] paths = Directory.GetFiles(directories[i], "*", SearchOption.AllDirectories);
                for (int j = 0; j < paths.Length; j++)
                {
                    try
                    {
                        FileInfo info = new FileInfo(paths[j]);
                        files.Add(new BackupFileEntry(paths[j], info.Length, info.LastWriteTimeUtc));
                    }
                    catch
                    {
                    }
                }
            }
            return files;
        }

        private static int RemoveEmptyDateFolders(string root)
        {
            int removed = 0;
            if (!Directory.Exists(root))
            {
                return 0;
            }

            string[] directories = Directory.GetDirectories(root);
            for (int i = 0; i < directories.Length; i++)
            {
                string name = Path.GetFileName(directories[i]);
                DateTime folderDate;
                if (!DateTime.TryParseExact(name, "yyyy-MM-dd", CultureInfo.InvariantCulture, DateTimeStyles.None, out folderDate))
                {
                    continue;
                }

                try
                {
                    if (Directory.GetFileSystemEntries(directories[i]).Length == 0)
                    {
                        Directory.Delete(directories[i], false);
                        removed++;
                    }
                }
                catch
                {
                }
            }
            return removed;
        }

        private static bool IsPathUnderRoot(string path, string root)
        {
            try
            {
                string fullPath = Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                string fullRoot = Path.GetFullPath(root).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                return fullPath.Equals(fullRoot, StringComparison.OrdinalIgnoreCase)
                    || fullPath.StartsWith(fullRoot + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase)
                    || fullPath.StartsWith(fullRoot + Path.AltDirectorySeparatorChar, StringComparison.OrdinalIgnoreCase);
            }
            catch
            {
                return false;
            }
        }

        private void LoadIndex()
        {
            try
            {
                _backedUpFingerprints.Clear();
                _fingerprintSavedPaths.Clear();
                _contentHashSavedPaths.Clear();
                _recentNameSizeEntries.Clear();

                if (!File.Exists(IndexFilePath))
                {
                    return;
                }

                string[] lines = File.ReadAllLines(IndexFilePath, Encoding.UTF8);
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i];
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }
                    string[] parts = line.Split('\t');
                    BackupRecord record = ParseBackupRecord(line);
                    if (record != null && record.HasBackupFiles && parts.Length > 0 && parts[0].Length > 0)
                    {
                        _backedUpFingerprints.Add(parts[0]);
                        if (parts.Length >= 6)
                        {
                            _fingerprintSavedPaths[parts[0]] = parts[5];
                        }
                        if (!string.IsNullOrEmpty(record.ContentHash))
                        {
                            _contentHashSavedPaths[record.ContentHash] = record.SavedPaths;
                        }
                        AddRecentNameSizeEntry(record.Kind, record.FileName, record.SizeBytes, record.Time, record.SavedPaths, record.Fingerprint);
                    }
                }
            }
            catch (Exception ex)
            {
                Log("读取索引失败：" + ex.Message);
            }
        }

        private void AppendIndex(string fingerprint, string sourcePath, FileInfo sourceInfo, string kind, string contentHash, List<string> saved)
        {
            try
            {
                EnsureLogDirectory();
                string line = string.Join("\t", new string[]
                {
                    fingerprint,
                    DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"),
                    sourceInfo.Length.ToString(),
                    sourceInfo.LastWriteTimeUtc.Ticks.ToString(),
                    sourcePath.Replace('\t', ' '),
                    string.Join(" ; ", saved.ToArray()).Replace('\t', ' '),
                    kind,
                    contentHash ?? ""
                });
                File.AppendAllText(IndexFilePath, line + Environment.NewLine, Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Log("写入索引失败：" + ex.Message);
            }
        }

        private bool IsFingerprintStillBackedUp(string fingerprint)
        {
            if (!_backedUpFingerprints.Contains(fingerprint))
            {
                return false;
            }

            string savedPaths;
            if (!_fingerprintSavedPaths.TryGetValue(fingerprint, out savedPaths) || string.IsNullOrWhiteSpace(savedPaths))
            {
                return true;
            }

            string[] paths = SplitSavedPaths(savedPaths);
            for (int i = 0; i < paths.Length; i++)
            {
                if (!string.IsNullOrWhiteSpace(paths[i]) && File.Exists(paths[i]))
                {
                    return true;
                }
            }

            _backedUpFingerprints.Remove(fingerprint);
            _fingerprintSavedPaths.Remove(fingerprint);
            RemoveFingerprintsFromIndex(new string[] { fingerprint });
            Log("检测到备份文件已不存在，已自动移除旧记录，允许重新备份。");
            return false;
        }

        private bool IsContentHashStillBackedUp(string contentHash)
        {
            string savedPaths;
            if (!_contentHashSavedPaths.TryGetValue(contentHash, out savedPaths))
            {
                return false;
            }

            if (AnySavedPathExists(savedPaths))
            {
                return true;
            }

            _contentHashSavedPaths.Remove(contentHash);
            return false;
        }

        private bool IsRecentNameSizeStillBackedUp(string kind, string fileName, long sizeBytes)
        {
            string key = MakeNameSizeKey(kind, fileName, sizeBytes);
            RecentNameSizeEntry entry;
            if (!_recentNameSizeEntries.TryGetValue(key, out entry))
            {
                return false;
            }

            if ((DateTime.Now - entry.Time).TotalMinutes > 10)
            {
                return false;
            }

            if (AnySavedPathExists(entry.SavedPaths))
            {
                return true;
            }

            _recentNameSizeEntries.Remove(key);
            return false;
        }

        private void AddRecentNameSizeEntry(string kind, string fileName, long sizeBytes, DateTime time, string savedPaths, string fingerprint)
        {
            string key = MakeNameSizeKey(kind, fileName, sizeBytes);
            RecentNameSizeEntry existing;
            if (_recentNameSizeEntries.TryGetValue(key, out existing) && existing.Time >= time)
            {
                return;
            }

            RecentNameSizeEntry entry = new RecentNameSizeEntry();
            entry.Time = time;
            entry.SavedPaths = savedPaths;
            entry.Fingerprint = fingerprint;
            _recentNameSizeEntries[key] = entry;
        }

        private static string MakeNameSizeKey(string kind, string fileName, long sizeBytes)
        {
            return (kind ?? "").Trim().ToLowerInvariant() + "|" + (fileName ?? "").Trim().ToLowerInvariant() + "|" + sizeBytes.ToString();
        }

        private static bool AnySavedPathExists(string savedPaths)
        {
            string[] paths = SplitSavedPaths(savedPaths);
            for (int i = 0; i < paths.Length; i++)
            {
                if (!string.IsNullOrWhiteSpace(paths[i]) && File.Exists(paths[i]))
                {
                    return true;
                }
            }
            return false;
        }

        private void RemoveFingerprintsFromIndex(string[] fingerprints)
        {
            try
            {
                if (!File.Exists(IndexFilePath) || fingerprints == null || fingerprints.Length == 0)
                {
                    return;
                }

                HashSet<string> removeSet = new HashSet<string>(fingerprints, StringComparer.OrdinalIgnoreCase);
                string[] lines = File.ReadAllLines(IndexFilePath, Encoding.UTF8);
                List<string> kept = new List<string>();
                for (int i = 0; i < lines.Length; i++)
                {
                    string line = lines[i];
                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    string[] parts = line.Split('\t');
                    if (parts.Length > 0 && removeSet.Contains(parts[0]))
                    {
                        continue;
                    }
                    kept.Add(line);
                }

                File.WriteAllLines(IndexFilePath, kept.ToArray(), Encoding.UTF8);
            }
            catch (Exception ex)
            {
                Log("更新备份记录失败：" + ex.Message);
            }
        }

        private BackupRecord ParseBackupRecord(string line)
        {
            if (string.IsNullOrWhiteSpace(line))
            {
                return null;
            }

            string[] parts = line.Split('\t');
            if (parts.Length < 6)
            {
                return null;
            }

            long size = 0;
            long.TryParse(parts[2], out size);

            DateTime timestamp;
            if (!DateTime.TryParse(parts[1], out timestamp))
            {
                timestamp = DateTime.MinValue;
            }

            string sourcePath = parts[4];
            string savedPaths = parts[5];
            string kind = parts.Length >= 7 && !string.IsNullOrEmpty(parts[6]) ? parts[6] : KindFromExtension(sourcePath);
            string contentHash = parts.Length >= 8 ? parts[7] : "";

            return new BackupRecord(
                parts[0],
                contentHash,
                timestamp,
                kind,
                Path.GetFileName(sourcePath),
                FormatSize(size),
                size,
                sourcePath,
                savedPaths,
                FirstSavedPath(savedPaths)
            );
        }

        private string UploadRawToCloudflare(string path, UploadConfig config)
        {
            PrepareTls();

            FileInfo info = new FileInfo(path);
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(config.CloudflareUploadUrl);
            request.Method = "POST";
            request.UserAgent = "ZeBackupAssistant/1.3";
            request.Timeout = 300000;
            request.ReadWriteTimeout = 300000;
            request.ContentType = GetContentType(path);
            request.Headers["X-Upload-Token"] = config.UploadToken;
            request.Headers["X-File-Name"] = Uri.EscapeDataString(info.Name);
            request.ContentLength = info.Length;

            using (Stream requestStream = request.GetRequestStream())
            {
                using (FileStream fileStream = File.OpenRead(path))
                {
                    CopyStream(fileStream, requestStream);
                }
            }

            string responseText = ReadWebResponseText(request);
            string link = ExtractJsonValue(responseText, "link");
            if (string.IsNullOrEmpty(link))
            {
                link = ExtractJsonValue(responseText, "url");
            }
            if (string.IsNullOrEmpty(link))
            {
                throw new InvalidOperationException("Cloudflare 上传页没有返回链接：" + TrimForMessage(responseText));
            }

            return link;
        }

        private UploadConfig LoadUploadConfig()
        {
            string configPath = FindExistingUploadConfigPath();
            if (string.IsNullOrEmpty(configPath))
            {
                EnsureCloudflareSampleConfig();
                configPath = UploadConfigPath;
            }

            Dictionary<string, string> values = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            string[] lines = File.ReadAllLines(configPath, Encoding.UTF8);
            for (int i = 0; i < lines.Length; i++)
            {
                string line = lines[i].Trim();
                if (line.Length == 0 || line.StartsWith("#"))
                {
                    continue;
                }

                int equals = line.IndexOf('=');
                if (equals <= 0)
                {
                    continue;
                }

                string key = line.Substring(0, equals).Trim();
                string value = line.Substring(equals + 1).Trim();
                values[key] = value;
            }

            string uploadUrl = GetConfigValue(values, "CloudflareUploadUrl");
            string token = GetConfigValue(values, "UploadToken");
            string adminToken = GetConfigValue(values, "AdminToken");
            if (!Uri.IsWellFormedUriString(uploadUrl, UriKind.Absolute) || uploadUrl.IndexOf("你的", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                throw new InvalidOperationException("还没配置 Cloudflare 上传地址。配置文件：" + configPath + "，把 pages.dev 的 /api/upload 地址填进去。");
            }
            if (string.IsNullOrWhiteSpace(token) || token.IndexOf("换成", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                throw new InvalidOperationException("还没配置 Cloudflare 上传口令。打开 " + configPath + "，把 UploadToken 换成你在 Cloudflare 设置的 UPLOAD_TOKEN。");
            }
            if (string.IsNullOrWhiteSpace(adminToken) || adminToken.IndexOf("换成", StringComparison.OrdinalIgnoreCase) >= 0)
            {
                adminToken = token;
            }

            return new UploadConfig(uploadUrl, token, adminToken);
        }

        private string FindExistingUploadConfigPath()
        {
            string[] paths = GetUploadConfigCandidatePaths();
            for (int i = 0; i < paths.Length; i++)
            {
                if (File.Exists(paths[i]))
                {
                    return paths[i];
                }
            }
            return "";
        }

        private string[] GetUploadConfigCandidatePaths()
        {
            List<string> paths = new List<string>();

            try
            {
                string exeDir = Path.GetDirectoryName(Application.ExecutablePath);
                if (!string.IsNullOrEmpty(exeDir))
                {
                    AddUniqueRoot(paths, Path.Combine(exeDir, UploadConfigFileName));
                }
            }
            catch
            {
            }

            AddUniqueRoot(paths, UploadConfigPath);

            string[] roots = GetAllBackupRoots();
            for (int i = 0; i < roots.Length; i++)
            {
                AddUniqueRoot(paths, Path.Combine(roots[i], "日志", UploadConfigFileName));
            }

            return paths.ToArray();
        }

        private void EnsureCloudflareSampleConfig()
        {
            EnsureLogDirectory();
            if (!string.IsNullOrEmpty(FindExistingUploadConfigPath()))
            {
                return;
            }

            string text =
                "# 泽PPT备份助手 Cloudflare 上传配置\r\n" +
                "# 当前推荐部署 D:\\安卓软件\\泽宁ppt\\cloudflare-pages-upload 里的 Pages/KV 上传页。\r\n" +
                "# 地址格式：https://你的项目名.pages.dev/api/upload\r\n" +
                "CloudflareUploadUrl=https://你的项目名.pages.dev/api/upload\r\n" +
                "\r\n" +
                "# 这里要和 Cloudflare Pages 里设置的 UPLOAD_TOKEN secret 完全一致。\r\n" +
                "UploadToken=换成你设置的UPLOAD_TOKEN\r\n" +
                "\r\n" +
                "# 这里要和 Cloudflare Pages 里设置的 ADMIN_TOKEN secret 完全一致，只给自己后台管理用。\r\n" +
                "AdminToken=换成你设置的ADMIN_TOKEN\r\n";
            File.WriteAllText(UploadConfigPath, text, Encoding.UTF8);
        }

        private static string GetConfigValue(Dictionary<string, string> values, string key)
        {
            string value;
            if (values.TryGetValue(key, out value))
            {
                return value;
            }
            return "";
        }

        private static void PrepareTls()
        {
            ServicePointManager.SecurityProtocol = ServicePointManager.SecurityProtocol | (SecurityProtocolType)3072;
        }

        private static string ReadWebResponseText(HttpWebRequest request)
        {
            try
            {
                using (HttpWebResponse response = (HttpWebResponse)request.GetResponse())
                {
                    return ReadStreamText(response.GetResponseStream());
                }
            }
            catch (WebException ex)
            {
                string status = "";
                string body = "";
                HttpWebResponse response = ex.Response as HttpWebResponse;
                if (response != null)
                {
                    status = "HTTP " + ((int)response.StatusCode).ToString() + " " + response.StatusDescription + "。";
                    body = ReadStreamText(response.GetResponseStream());
                    response.Close();
                }
                throw new InvalidOperationException(status + TrimForMessage(body.Length > 0 ? body : ex.Message));
            }
        }

        private static string ReadStreamText(Stream stream)
        {
            if (stream == null)
            {
                return "";
            }

            using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
            {
                return reader.ReadToEnd();
            }
        }

        private static string ExtractJsonValue(string json, string name)
        {
            if (string.IsNullOrEmpty(json))
            {
                return "";
            }

            Match match = Regex.Match(json, "\"" + Regex.Escape(name) + "\"\\s*:\\s*\"((?:\\\\.|[^\"])*)\"", RegexOptions.IgnoreCase);
            if (!match.Success)
            {
                return "";
            }

            try
            {
                return Regex.Unescape(match.Groups[1].Value);
            }
            catch
            {
                return match.Groups[1].Value;
            }
        }

        private static string TrimForMessage(string value)
        {
            if (string.IsNullOrEmpty(value))
            {
                return "";
            }

            value = value.Replace("\r", " ").Replace("\n", " ").Trim();
            if (value.Length > 300)
            {
                return value.Substring(0, 300) + "...";
            }
            return value;
        }

        private static void CopyStream(Stream source, Stream target)
        {
            byte[] buffer = new byte[64 * 1024];
            int read;
            while ((read = source.Read(buffer, 0, buffer.Length)) > 0)
            {
                target.Write(buffer, 0, read);
            }
        }

        private static string GetContentType(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".pdf")
            {
                return "application/pdf";
            }
            if (ext == ".ppt")
            {
                return "application/vnd.ms-powerpoint";
            }
            if (ext == ".pptx")
            {
                return "application/vnd.openxmlformats-officedocument.presentationml.presentation";
            }
            if (ext == ".pps")
            {
                return "application/vnd.ms-powerpoint";
            }
            if (ext == ".ppsx")
            {
                return "application/vnd.openxmlformats-officedocument.presentationml.slideshow";
            }
            if (ext == ".doc")
            {
                return "application/msword";
            }
            if (ext == ".docx")
            {
                return "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            }
            if (ext == ".rtf")
            {
                return "application/rtf";
            }
            return "application/octet-stream";
        }

        private static string[] SplitSavedPaths(string savedPaths)
        {
            if (string.IsNullOrWhiteSpace(savedPaths))
            {
                return new string[0];
            }

            return savedPaths.Split(new string[] { " ; " }, StringSplitOptions.RemoveEmptyEntries);
        }

        private static string KindFromExtension(string path)
        {
            string ext = Path.GetExtension(path).ToLowerInvariant();
            if (ext == ".pdf")
            {
                return "PDF";
            }
            if (ext == ".doc" || ext == ".docx" || ext == ".docm" || ext == ".rtf" || ext == ".wps")
            {
                return "Word";
            }
            return "PPT";
        }

        private bool IsKindEnabled(string kind)
        {
            if (string.Equals(kind, "PDF", StringComparison.OrdinalIgnoreCase))
            {
                return _backupPdfEnabled;
            }
            if (string.Equals(kind, "Word", StringComparison.OrdinalIgnoreCase))
            {
                return _backupWordEnabled;
            }
            return _backupPptEnabled;
        }

        private static string FirstSavedPath(string savedPaths)
        {
            if (string.IsNullOrEmpty(savedPaths))
            {
                return "";
            }

            string[] paths = savedPaths.Split(new string[] { " ; " }, StringSplitOptions.None);
            if (paths.Length == 0)
            {
                return savedPaths.Trim();
            }
            return paths[0].Trim();
        }

        public static string FormatSize(long bytes)
        {
            if (bytes >= 1024L * 1024L * 1024L)
            {
                return (bytes / 1024d / 1024d / 1024d).ToString("0.##") + " GB";
            }
            if (bytes >= 1024L * 1024L)
            {
                return (bytes / 1024d / 1024d).ToString("0.##") + " MB";
            }
            if (bytes >= 1024L)
            {
                return (bytes / 1024d).ToString("0.##") + " KB";
            }
            return bytes.ToString() + " B";
        }

        private void Log(string message)
        {
            string line = "[" + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + "] " + message + Environment.NewLine;

            try
            {
                EnsureLogDirectory();
                RotateLogIfNeeded();
                File.AppendAllText(LogFilePath, line, Encoding.UTF8);
                return;
            }
            catch
            {
            }

            try
            {
                string fallback = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "泽.log");
                File.AppendAllText(fallback, line, Encoding.UTF8);
            }
            catch
            {
            }
        }

        private void RotateLogIfNeeded()
        {
            try
            {
                if (!File.Exists(LogFilePath))
                {
                    return;
                }

                FileInfo info = new FileInfo(LogFilePath);
                if (info.Length < 5 * 1024 * 1024)
                {
                    return;
                }

                string archive = Path.Combine(Path.GetDirectoryName(LogFilePath), "泽-" + DateTime.Now.ToString("yyyyMMdd-HHmmss") + ".log");
                File.Move(LogFilePath, archive);
            }
            catch
            {
            }
        }

        private static bool IsAllowedDocument(string path)
        {
            try
            {
                string ext = Path.GetExtension(path);
                for (int i = 0; i < AllowedExtensions.Length; i++)
                {
                    if (string.Equals(ext, AllowedExtensions[i], StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
            catch
            {
            }
            return false;
        }

        private static bool IsInsideBackupRoot(string path)
        {
            try
            {
                string fullPath = Path.GetFullPath(path).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                string[] roots = GetAllBackupRoots();
                for (int i = 0; i < roots.Length; i++)
                {
                    string root = Path.GetFullPath(roots[i]).TrimEnd(Path.DirectorySeparatorChar, Path.AltDirectorySeparatorChar);
                    if (fullPath.Equals(root, StringComparison.OrdinalIgnoreCase) ||
                        fullPath.StartsWith(root + Path.DirectorySeparatorChar, StringComparison.OrdinalIgnoreCase) ||
                        fullPath.StartsWith(root + Path.AltDirectorySeparatorChar, StringComparison.OrdinalIgnoreCase))
                    {
                        return true;
                    }
                }
            }
            catch
            {
            }
            return false;
        }

        private static string ComputeContentHash(string path)
        {
            try
            {
                using (FileStream stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
                using (SHA256 sha = SHA256.Create())
                {
                    byte[] hash = sha.ComputeHash(stream);
                    return BytesToHex(hash);
                }
            }
            catch
            {
                return "";
            }
        }

        private static string MakeFingerprint(string sourcePath, long length, long lastWriteUtcTicks)
        {
            string normalized = sourcePath.Trim().ToLowerInvariant() + "|" + length.ToString() + "|" + lastWriteUtcTicks.ToString();
            byte[] bytes = Encoding.UTF8.GetBytes(normalized);
            using (SHA256 sha = SHA256.Create())
            {
                byte[] hash = sha.ComputeHash(bytes);
                return BytesToHex(hash);
            }
        }

        private static string BytesToHex(byte[] hash)
        {
            StringBuilder builder = new StringBuilder(hash.Length * 2);
            for (int i = 0; i < hash.Length; i++)
            {
                builder.Append(hash[i].ToString("x2"));
            }
            return builder.ToString();
        }

        private void RaiseBackupSucceeded(string fileName, string savedPaths)
        {
            EventHandler<BackupCompletedEventArgs> handler = BackupSucceeded;
            if (handler == null)
            {
                return;
            }

            try
            {
                handler(this, new BackupCompletedEventArgs(fileName, savedPaths));
            }
            catch
            {
            }
        }
    }

    internal sealed class BackupCompletedEventArgs : EventArgs
    {
        public readonly string FileName;
        public readonly string SavedPaths;

        public BackupCompletedEventArgs(string fileName, string savedPaths)
        {
            FileName = fileName;
            SavedPaths = savedPaths;
        }
    }

    internal sealed class CleanupResult
    {
        public int DeletedFiles;
        public long FreedBytes;
        public int RemovedFolders;
    }

    internal sealed class BackupFileEntry
    {
        public readonly string Path;
        public readonly long SizeBytes;
        public readonly DateTime LastWriteUtc;

        public BackupFileEntry(string path, long sizeBytes, DateTime lastWriteUtc)
        {
            Path = path;
            SizeBytes = sizeBytes;
            LastWriteUtc = lastWriteUtc;
        }
    }

    internal sealed class PendingDocument
    {
        public string Path;
        public long LastLength;
        public long LastWriteUtcTicks;
        public DateTime StableSinceUtc;
    }

    internal sealed class SearchDir
    {
        public readonly string Path;
        public readonly int Depth;

        public SearchDir(string path, int depth)
        {
            Path = path;
            Depth = depth;
        }
    }

    internal sealed class TitleSearchCacheEntry
    {
        public DateTime CheckedUtc;
        public string Path;
    }

    internal sealed class RecentNameSizeEntry
    {
        public DateTime Time;
        public string SavedPaths;
        public string Fingerprint;
    }

    internal sealed class OpenDocument
    {
        public readonly string Path;
        public readonly string Kind;
        public readonly string AppName;

        public OpenDocument(string path, string kind, string appName)
        {
            Path = path;
            Kind = kind;
            AppName = appName;
        }
    }

    internal sealed class UploadConfig
    {
        public readonly string CloudflareUploadUrl;
        public readonly string UploadToken;
        public readonly string AdminToken;

        public UploadConfig(string cloudflareUploadUrl, string uploadToken, string adminToken)
        {
            CloudflareUploadUrl = cloudflareUploadUrl;
            UploadToken = uploadToken;
            AdminToken = adminToken;
        }
    }

    internal sealed class BackupRecord
    {
        public readonly string Fingerprint;
        public readonly string ContentHash;
        public readonly DateTime Time;
        public readonly string Kind;
        public readonly string FileName;
        public readonly string SizeText;
        public readonly long SizeBytes;
        public readonly string SourcePath;
        public readonly string SavedPaths;
        public readonly string PrimaryTarget;

        public BackupRecord(string fingerprint, string contentHash, DateTime time, string kind, string fileName, string sizeText, long sizeBytes, string sourcePath, string savedPaths, string primaryTarget)
        {
            Fingerprint = fingerprint;
            ContentHash = contentHash;
            Time = time;
            Kind = kind;
            FileName = fileName;
            SizeText = sizeText;
            SizeBytes = sizeBytes;
            SourcePath = sourcePath;
            SavedPaths = savedPaths;
            PrimaryTarget = primaryTarget;
        }

        public string TimeText
        {
            get
            {
                if (Time == DateTime.MinValue)
                {
                    return "";
                }
                return Time.ToString("yyyy-MM-dd HH:mm:ss");
            }
        }

        public bool HasBackupFiles
        {
            get
            {
                string[] paths = SavedPaths.Split(new string[] { " ; " }, StringSplitOptions.RemoveEmptyEntries);
                for (int i = 0; i < paths.Length; i++)
                {
                    if (!string.IsNullOrWhiteSpace(paths[i]) && File.Exists(paths[i]))
                    {
                        return true;
                    }
                }
                return false;
            }
        }

        public string BackupStatusText
        {
            get { return HasBackupFiles ? "存在" : "已删除"; }
        }
    }
}
