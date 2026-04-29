using System;
using System.Threading.Tasks;
using Microsoft.Win32;

namespace Office365CleanupTool.Services
{
    public enum NetworkRepairStatus
    {
        Processing,
        Success,
        Warning
    }

    public class NetworkRepairProgressEventArgs : EventArgs
    {
        public NetworkRepairProgressEventArgs(string message, NetworkRepairStatus status)
        {
            Message = message;
            Status = status;
        }

        public string Message { get; }

        public NetworkRepairStatus Status { get; }
    }

    public class NetworkRepairService
    {
        private readonly IScriptRunner _scriptRunner;

        public NetworkRepairService(IScriptRunner scriptRunner)
        {
            _scriptRunner = scriptRunner;
        }

        public event EventHandler<NetworkRepairProgressEventArgs>? ProgressChanged;

        public async Task RunAsync(UiLanguage language = UiLanguage.Chinese)
        {
            OnProgressChanged(T(language, "正在修复 Office 激活网络问题...", "Repairing Office activation network issues..."), NetworkRepairStatus.Processing);

            await Task.Run(() =>
            {
                UpdateRegistrySettings(language);
                ImportProxyFromInternetSettings(language);
            });
        }

        private void UpdateRegistrySettings(UiLanguage language)
        {
            OnProgressChanged(T(language, "正在修改注册表设置...", "Updating registry settings..."), NetworkRepairStatus.Processing);

            try
            {
                using (var key = Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Internet Settings", true))
                {
                    if (key == null)
                    {
                        throw new Exception(T(language, "无法打开注册表项", "Cannot open registry key"));
                    }

                    key.SetValue("SecureProtocols", 2720, RegistryValueKind.DWord);
                    OnProgressChanged(T(language, "✓ SecureProtocols 设置成功", "✓ SecureProtocols updated"), NetworkRepairStatus.Success);

                    key.SetValue("AutoDetect", 1, RegistryValueKind.DWord);
                    OnProgressChanged(T(language, "✓ AutoDetect 设置成功", "✓ AutoDetect updated"), NetworkRepairStatus.Success);

                    key.SetValue("AutoConfigURL", string.Empty, RegistryValueKind.String);
                    OnProgressChanged(T(language, "✓ AutoConfigURL 设置成功", "✓ AutoConfigURL updated"), NetworkRepairStatus.Success);
                }
            }
            catch (Exception ex)
            {
                throw new Exception(T(language, $"注册表修改失败: {ex.Message}", $"Failed to update registry settings: {ex.Message}"), ex);
            }
        }

        private void ImportProxyFromInternetSettings(UiLanguage language)
        {
            OnProgressChanged(T(language, "正在设置网络代理...", "Importing network proxy settings..."), NetworkRepairStatus.Processing);

            var executionResult = _scriptRunner.RunAsync(new ScriptExecutionRequest
            {
                FileName = "netsh",
                Arguments = "winhttp import proxy source=ie"
            }).GetAwaiter().GetResult();

            if (!executionResult.IsSuccess)
            {
                OnProgressChanged(T(language, $"网络代理设置警告: {executionResult.StandardError}", $"Proxy import warning: {executionResult.StandardError}"), NetworkRepairStatus.Warning);
                return;
            }

            OnProgressChanged(T(language, "✓ 网络代理设置成功", "✓ Network proxy imported"), NetworkRepairStatus.Success);
        }

        private void OnProgressChanged(string message, NetworkRepairStatus status)
        {
            ProgressChanged?.Invoke(this, new NetworkRepairProgressEventArgs(message, status));
        }

        private static string T(UiLanguage language, string zh, string en)
        {
            return LocalizationService.Localize(language, zh, en);
        }
    }
}
