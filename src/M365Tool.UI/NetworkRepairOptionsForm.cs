using Office365CleanupTool.Services;

namespace Office365CleanupTool;

internal sealed class NetworkRepairOptionsForm : Form
{
    public NetworkRepairOptions Options { get; } = new();

    public NetworkRepairOptionsForm(NetworkRepairService service, UiLanguage language)
    {
        string T(string zh, string en) => LocalizationService.Localize(language, zh, en);
        Text = T("选择网络修复项目", "Select Network Repairs");
        ClientSize = new Size(640, 390);
        StartPosition = FormStartPosition.CenterParent;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;
        MinimizeBox = false;
        Font = WorkbenchUi.CreateUiFont(10F);
        var tls = new CheckBox { Text = T("启用 TLS 1.2，保留现有协议设置", "Enable TLS 1.2 and preserve existing protocol settings"), Checked = true, Bounds = new Rectangle(20, 18, 600, 28) };
        var detect = new CheckBox { Text = T("启用自动代理检测", "Enable automatic proxy detection"), Bounds = new Rectangle(20, 54, 600, 28) };
        var proxy = new CheckBox { Text = T("用当前用户代理替换 WinHTTP 代理", "Replace WinHTTP proxy with current-user proxy"), Bounds = new Rectangle(20, 90, 600, 28) };
        var preview = new TextBox { Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical, Bounds = new Rectangle(20, 132, 600, 176) };
        var apply = new Button { Text = T("备份并修复", "Back Up and Repair"), DialogResult = DialogResult.OK, Bounds = new Rectangle(300, 330, 180, 38) };
        var cancel = new Button { Text = T("取消", "Cancel"), DialogResult = DialogResult.Cancel, Bounds = new Rectangle(490, 330, 130, 38) };
        void RefreshPreview()
        {
            Options.EnableTls12 = tls.Checked;
            Options.EnableAutoDetect = detect.Checked;
            Options.ImportUserProxy = proxy.Checked;
            apply.Enabled = tls.Checked || detect.Checked || proxy.Checked;
            preview.Text = service.GetRepairPreview(Options, language);
        }
        tls.CheckedChanged += (_, _) => RefreshPreview();
        detect.CheckedChanged += (_, _) => RefreshPreview();
        proxy.CheckedChanged += (_, _) => RefreshPreview();
        Controls.AddRange(new Control[] { tls, detect, proxy, preview, apply, cancel });
        AcceptButton = apply;
        CancelButton = cancel;
        RefreshPreview();
    }
}
