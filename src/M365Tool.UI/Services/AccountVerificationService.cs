using System;
using System.Net.Http;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace Office365CleanupTool.Services
{
    public enum AccountCloudKind
    {
        Unknown,
        China21V,
        Global
    }

    public sealed class AccountVerificationResult
    {
        public bool IsSuccess { get; set; }

        public bool Is21VAccount { get; set; }

        public AccountCloudKind CloudKind { get; set; }

        public string Account { get; set; } = string.Empty;

        public string Diagnostic { get; set; } = string.Empty;
    }

    public sealed class AccountVerificationService
    {
        private static readonly HttpClient HttpClient = new()
        {
            Timeout = TimeSpan.FromSeconds(12)
        };

        public async Task<AccountVerificationResult> Verify21VAccountAsync(
            string account,
            UiLanguage language = UiLanguage.Chinese)
        {
            string trimmed = (account ?? string.Empty).Trim();
            if (!Regex.IsMatch(trimmed, @"^[^@\s]+@[^@\s]+\.[^@\s]+$", RegexOptions.CultureInvariant))
            {
                return new AccountVerificationResult
                {
                    IsSuccess = false,
                    Is21VAccount = false,
                    CloudKind = AccountCloudKind.Unknown,
                    Account = trimmed,
                    Diagnostic = T(language, "账号格式无效。", "Invalid account format.")
                };
            }

            string domain = trimmed[(trimmed.IndexOf('@') + 1)..];
            string chinaUrl = $"https://login.partner.microsoftonline.cn/{domain}/v2.0/.well-known/openid-configuration";
            string globalUrl = $"https://login.microsoftonline.com/{domain}/v2.0/.well-known/openid-configuration";

            (bool chinaOk, string chinaBody, string chinaError) = await TryGetOpenIdAsync(chinaUrl);
            (bool globalOk, string globalBody, string globalError) = await TryGetOpenIdAsync(globalUrl);

            if (chinaOk && (chinaBody.Contains("microsoftonline.cn", StringComparison.OrdinalIgnoreCase) ||
                            chinaBody.Contains("chinacloudapi.cn", StringComparison.OrdinalIgnoreCase)))
            {
                return new AccountVerificationResult
                {
                    IsSuccess = true,
                    Is21VAccount = true,
                    CloudKind = AccountCloudKind.China21V,
                    Account = trimmed,
                    Diagnostic = T(language, "账号已通过 21V 云端点验证。", "Account validated against 21V cloud endpoint.")
                };
            }

            if (globalOk)
            {
                return new AccountVerificationResult
                {
                    IsSuccess = true,
                    Is21VAccount = false,
                    CloudKind = AccountCloudKind.Global,
                    Account = trimmed,
                    Diagnostic = T(language, "账号可在全球版端点解析，未识别为 21V 账号。", "Account resolved on global endpoint and is not recognized as 21V.")
                };
            }

            return new AccountVerificationResult
            {
                IsSuccess = false,
                Is21VAccount = false,
                CloudKind = AccountCloudKind.Unknown,
                Account = trimmed,
                Diagnostic = T(
                    language,
                    $"无法验证账号云环境。ChinaEndpoint={chinaError}; GlobalEndpoint={globalError}",
                    $"Unable to verify account cloud environment. ChinaEndpoint={chinaError}; GlobalEndpoint={globalError}")
            };
        }

        private static async Task<(bool Success, string Body, string Error)> TryGetOpenIdAsync(string url)
        {
            try
            {
                using var request = new HttpRequestMessage(HttpMethod.Get, url);
                request.Headers.UserAgent.ParseAdd("M365Tool/1.0");
                using HttpResponseMessage response = await HttpClient.SendAsync(request);
                if (!response.IsSuccessStatusCode)
                {
                    return (false, string.Empty, $"{(int)response.StatusCode}");
                }

                string body = await response.Content.ReadAsStringAsync();
                return (true, body, string.Empty);
            }
            catch (Exception ex)
            {
                return (false, string.Empty, ex.Message);
            }
        }

        private static string T(UiLanguage language, string zh, string en)
        {
            return LocalizationService.Localize(language, zh, en);
        }
    }
}
