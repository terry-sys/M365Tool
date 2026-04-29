using System;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;

namespace Office365CleanupTool.Services
{
    public class CleanupProgressEventArgs : EventArgs
    {
        public CleanupProgressEventArgs(string message, int progress)
        {
            Message = message;
            Progress = progress;
        }

        public string Message { get; }

        public int Progress { get; }
    }

    public class CleanupService
    {
        private const string SignOutScriptFileName = "signoutofwamaccounts.ps1";
        private const string OrganizationClientId = "d3590ed6-52b3-4102-aeff-aad2292ab01c";

        private static readonly string[] CleanupUrls =
        {
            "https://download.microsoft.com/download/e/1/b/e1bbdc16-fad4-4aa2-a309-2ba3cae8d424/OLicenseCleanup.zip",
            "https://download.microsoft.com/download/f/8/7/f8745d3b-49ad-4eac-b49a-2fa60b929e7d/signoutofwamaccounts.zip",
            "https://download.microsoft.com/download/8/e/f/8ef13ae0-6aa8-48a2-8697-5b1711134730/WPJCleanUp.zip"
        };

        private readonly IScriptRunner _scriptRunner;

        public CleanupService(IScriptRunner scriptRunner)
        {
            _scriptRunner = scriptRunner;
        }

        public event EventHandler<CleanupProgressEventArgs>? ProgressChanged;

        public async Task RunAsync(string extractPath, UiLanguage language = UiLanguage.Chinese)
        {
            OnProgressChanged(T(language, "开始清理流程...", "Starting cleanup flow..."), 0);

            OnProgressChanged(T(language, "创建临时目录...", "Creating temp directory..."), 10);
            Directory.CreateDirectory(extractPath);

            for (int i = 0; i < CleanupUrls.Length; i++)
            {
                string url = CleanupUrls[i];
                string fileName = Path.GetFileName(url);
                string downloadPath = Path.Combine(extractPath, fileName);

                OnProgressChanged(T(language, $"下载文件 {i + 1}/{CleanupUrls.Length}: {fileName}", $"Downloading file {i + 1}/{CleanupUrls.Length}: {fileName}"), 20 + (i * 15));
                await DownloadFileAsync(url, downloadPath);

                OnProgressChanged(T(language, $"解压文件: {fileName}", $"Extracting file: {fileName}"), 35 + (i * 15));
                await ExtractFileAsync(downloadPath, extractPath);
            }

            OnProgressChanged(T(language, "验证关键文件...", "Validating required files..."), 70);
            ValidateFiles(extractPath, language);

            OnProgressChanged(T(language, "修复 signoutofwamaccounts.ps1 兼容性...", "Applying signoutofwamaccounts.ps1 compatibility fix..."), 75);
            NormalizeSignOutScript(extractPath);

            OnProgressChanged(T(language, "执行 OLicenseCleanup.vbs...", "Running OLicenseCleanup.vbs..."), 80);
            await ExecuteScriptAsync(language, "cscript.exe", Path.Combine(extractPath, "OLicenseCleanup", "OLicenseCleanup.vbs"));

            OnProgressChanged(T(language, "执行 signoutofwamaccounts.ps1...", "Running signoutofwamaccounts.ps1..."), 90);
            await ExecuteScriptAsync(
                language,
                "powershell.exe",
                "-ExecutionPolicy",
                "Unrestricted",
                "-File",
                Path.Combine(extractPath, SignOutScriptFileName));

            OnProgressChanged(T(language, "执行 WPJCleanUp.cmd...", "Running WPJCleanUp.cmd..."), 100);
            await ExecuteScriptAsync(language, "cmd.exe", "/C", Path.Combine(extractPath, @"WPJCleanUp\WPJCleanUp\WPJCleanUp.cmd"));
        }

        private async Task DownloadFileAsync(string url, string filePath)
        {
            using (var client = new HttpClient())
            using (var response = await client.GetAsync(url, HttpCompletionOption.ResponseHeadersRead))
            {
                response.EnsureSuccessStatusCode();
                using (var fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write, FileShare.None))
                {
                    await response.Content.CopyToAsync(fileStream);
                }
            }
        }

        private async Task ExtractFileAsync(string zipPath, string extractPath)
        {
            await Task.Run(() => ZipFile.ExtractToDirectory(zipPath, extractPath, true));
        }

        private void ValidateFiles(string extractPath, UiLanguage language)
        {
            string[] requiredFiles =
            {
                Path.Combine(extractPath, "OLicenseCleanup", "OLicenseCleanup.vbs"),
                Path.Combine(extractPath, SignOutScriptFileName),
                Path.Combine(extractPath, @"WPJCleanUp\WPJCleanUp\WPJCleanUp.cmd")
            };

            string[] fileNames =
            {
                "OLicenseCleanup.vbs",
                SignOutScriptFileName,
                "WPJCleanUp.cmd"
            };

            for (int i = 0; i < requiredFiles.Length; i++)
            {
                if (!File.Exists(requiredFiles[i]))
                {
                    throw new FileNotFoundException(
                        T(language, $"文件验证失败：{fileNames[i]} 不存在。", $"File validation failed: {fileNames[i]} does not exist."),
                        requiredFiles[i]);
                }
            }
        }

        private void NormalizeSignOutScript(string extractPath)
        {
            string scriptPath = Path.Combine(extractPath, SignOutScriptFileName);
            byte[] bytes = File.ReadAllBytes(scriptPath);

            for (int i = 0; i < bytes.Length; i++)
            {
                // 上游压缩包里的脚本常出现 ANSI 智能引号（0x93/0x94），
                // 在中文系统代码页下会被错误解码为乱码，导致 PowerShell 语法错误。
                if (bytes[i] == 0x93 || bytes[i] == 0x94)
                {
                    bytes[i] = 0x22; // "
                }
            }

            string script = Encoding.UTF8.GetString(bytes);
            string[] lines = script.Replace("\r\n", "\n").Split('\n');
            string canonicalLine = $"$accounts.Accounts | % {{ AwaitAction ($_.SignOutAsync(\"{OrganizationClientId}\")) }}";

            bool replaced = false;
            for (int i = 0; i < lines.Length; i++)
            {
                if (lines[i].TrimStart().StartsWith("$accounts.Accounts | %", StringComparison.Ordinal))
                {
                    lines[i] = canonicalLine;
                    replaced = true;
                    break;
                }
            }

            if (!replaced)
            {
                Array.Resize(ref lines, lines.Length + 1);
                lines[^1] = canonicalLine;
            }

            string normalizedScript = string.Join("\r\n", lines).TrimEnd() + "\r\n";
            File.WriteAllText(scriptPath, normalizedScript, new UTF8Encoding(encoderShouldEmitUTF8Identifier: true));
        }

        private async Task ExecuteScriptAsync(UiLanguage language, string fileName, params string[] arguments)
        {
            var result = await _scriptRunner.RunAsync(new ScriptExecutionRequest
            {
                FileName = fileName,
                Arguments = string.Join(" ", arguments)
            });

            if (!result.IsSuccess)
            {
                throw new Exception(T(
                    language,
                    $"脚本执行失败: {fileName}\n{result.StandardError}",
                    $"Script execution failed: {fileName}\n{result.StandardError}"));
            }
        }

        private void OnProgressChanged(string message, int progress)
        {
            ProgressChanged?.Invoke(this, new CleanupProgressEventArgs(message, progress));
        }

        private static string T(UiLanguage language, string zh, string en)
        {
            return LocalizationService.Localize(language, zh, en);
        }
    }
}
