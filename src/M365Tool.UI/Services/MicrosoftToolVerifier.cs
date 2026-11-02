using System.Text;

namespace Office365CleanupTool.Services;

internal static class MicrosoftToolVerifier
{
    internal static string BuildVerificationArguments(string path)
    {
        string literal = path.Replace("'", "''");
        string script = $"$ErrorActionPreference='Stop'; $s=Get-AuthenticodeSignature -LiteralPath '{literal}'; " +
            "if ($s.Status -ne 'Valid' -or $null -eq $s.SignerCertificate -or " +
            "$s.SignerCertificate.GetNameInfo([System.Security.Cryptography.X509Certificates.X509NameType]::SimpleName,$false) -ne 'Microsoft Corporation') " +
            "{ throw 'The tool does not have a valid Microsoft signature.' }; Write-Output $s.SignerCertificate.Subject";
        return "-NoProfile -NonInteractive -EncodedCommand " + Convert.ToBase64String(Encoding.Unicode.GetBytes(script));
    }

    public static async Task VerifyAsync(string path, IScriptRunner runner, CancellationToken cancellationToken)
    {
        var result = await runner.RunAsync(new ScriptExecutionRequest
        {
            FileName = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.System), "WindowsPowerShell", "v1.0", "powershell.exe"),
            Arguments = BuildVerificationArguments(path),
            CancellationToken = cancellationToken,
            Timeout = TimeSpan.FromMinutes(2),
            KillOnCancellation = true
        });
        cancellationToken.ThrowIfCancellationRequested();
        if (!result.IsSuccess)
        {
            throw new InvalidOperationException($"Microsoft signature verification failed: {path}\n{result.StandardError}");
        }
    }
}
