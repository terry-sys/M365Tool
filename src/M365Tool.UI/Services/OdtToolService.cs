using System.Net.Http;
using System.Text.RegularExpressions;

namespace Office365CleanupTool.Services;

internal sealed class OdtToolService(IScriptRunner runner)
{
    private const string OdtDetailsUrl = "https://www.microsoft.com/en-us/download/details.aspx?id=49117";

    public async Task<string> PrepareAsync(string directory, CancellationToken cancellationToken, Action<string> progress, UiLanguage language = UiLanguage.Chinese)
    {
        Directory.CreateDirectory(directory);
        string setup = Path.Combine(directory, "setup.exe");
        if (File.Exists(setup))
        {
            progress(LocalizationService.Localize(language, "正在验证 ODT 的微软数字签名...", "Verifying Microsoft ODT signature..."));
            await MicrosoftToolVerifier.VerifyAsync(setup, runner, cancellationToken);
            return setup;
        }

        string installer = Path.Combine(ToolPathService.GetOdtCacheDirectory(), "ODTSetup.exe");
        if (!File.Exists(installer))
        {
            progress(LocalizationService.Localize(language, "正在下载 ODT...", "Downloading Office Deployment Tool..."));
            string url = await GetOdtDownloadUrlAsync(cancellationToken);
            string partial = installer + "." + Guid.NewGuid().ToString("N") + ".exe";
            try
            {
                await ToolDownloadService.DownloadAsync(url, partial, cancellationToken);
                await MicrosoftToolVerifier.VerifyAsync(partial, runner, cancellationToken);
                File.Move(partial, installer, overwrite: true);
            }
            finally
            {
                if (File.Exists(partial)) File.Delete(partial);
            }
        }
        progress(LocalizationService.Localize(language, "正在验证 ODT 安装包签名...", "Verifying ODT package signature..."));
        await MicrosoftToolVerifier.VerifyAsync(installer, runner, cancellationToken);
        progress(LocalizationService.Localize(language, "正在解压 ODT...", "Extracting Office Deployment Tool..."));
        var result = await runner.RunAsync(new ScriptExecutionRequest
        {
            FileName = installer,
            Arguments = $"/quiet /extract:\"{directory}\"",
            WorkingDirectory = directory,
            CancellationToken = cancellationToken,
            Timeout = TimeSpan.FromMinutes(5),
            KillOnCancellation = true
        });
        cancellationToken.ThrowIfCancellationRequested();
        if (!result.IsSuccess || !File.Exists(setup))
            throw new InvalidOperationException("ODT extraction failed. " + result.StandardError);
        await MicrosoftToolVerifier.VerifyAsync(setup, runner, cancellationToken);
        return setup;
    }

    private static async Task<string> GetOdtDownloadUrlAsync(CancellationToken cancellationToken)
    {
        using var client = new HttpClient();
        client.DefaultRequestHeaders.UserAgent.ParseAdd("M365Tool/1.0");
        string pageContent = await client.GetStringAsync(OdtDetailsUrl, cancellationToken);

        string[] patterns =
        {
            "\"url\":\"(https:\\\\/\\\\/download\\\\.microsoft\\\\.com\\\\/[^\\\"]*officedeploymenttool[^\\\"]*)\"",
            "href=\"(https://download\\.microsoft\\.com/[^\"]*officedeploymenttool[^\"]*)\"",
            "href='(https://download\\.microsoft\\.com/[^']*officedeploymenttool[^']*)'",
            "(https://download\\.microsoft\\.com/\\S*officedeploymenttool\\S*\\.exe)"
        };

        foreach (string pattern in patterns)
        {
            var match = Regex.Match(pageContent, pattern, RegexOptions.IgnoreCase);
            if (match.Success)
            {
                return match.Groups[1].Value.Replace("\\/", "/");
            }
        }

        throw new InvalidOperationException("未能从微软官方下载页解析出 ODT 下载地址。");
    }

}
