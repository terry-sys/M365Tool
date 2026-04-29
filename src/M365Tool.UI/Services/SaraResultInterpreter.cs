using Office365CleanupTool.Models;

namespace Office365CleanupTool.Services
{
    public static class SaraResultInterpreter
    {
        public static string GetSummary(int exitCode, UiLanguage language = UiLanguage.Chinese)
        {
            return exitCode switch
            {
                0 => T(language, "卸载场景执行成功，建议重启计算机完成剩余清理。", "Uninstall scenario completed successfully. Restart is recommended to finish cleanup."),
                6 => T(language, "检测到 Office 程序仍在运行，请关闭后重试。", "Office processes are still running. Close them and retry."),
                8 => T(language, "检测到多个 Office 版本，请改为选择明确版本后重新卸载。", "Multiple Office versions were detected. Select a specific version and retry."),
                9 => T(language, "SaRA 未能完成卸载，建议查看日志后改用完整界面版 SaRA。", "SaRA could not complete uninstall. Check logs and consider full UI SaRA."),
                10 => T(language, "卸载需要管理员权限，请以管理员身份运行工具。", "Administrator privileges are required for uninstall."),
                66 => T(language, "指定的 Office 版本与当前设备检测结果不匹配。", "Selected Office version does not match detected install."),
                67 => T(language, "传入的 Office 版本参数无效。", "The Office version parameter is invalid."),
                68 => T(language, "未检测到可卸载的 Office 版本。", "No uninstallable Office version was detected."),
                _ => T(language, $"SaRA 返回了未映射的结果码: {exitCode}。", $"SaRA returned an unmapped code: {exitCode}.")
            };
        }

        public static string GetNextStepAdvice(
            int exitCode,
            OfficeUninstallTargetOption target,
            string detectedOfficeSummary,
            UiLanguage language = UiLanguage.Chinese)
        {
            return exitCode switch
            {
                0 => T(language, "重启计算机，然后检查是否还存在残留 Office 组件。如果仍有残留，我们下一步可以补做残留清理。", "Restart the computer, then check for remaining Office components."),
                6 => T(language, "请确认 Word、Excel、Outlook、Teams 等都已关闭，然后重新执行卸载。", "Close Word, Excel, Outlook, Teams, then retry uninstall."),
                8 => T(
                    language,
                    $"当前设备检测到多个 Office 条目。建议不要再用“{target.DisplayName}”自动模式，而是改选明确版本重新卸载。当前检测结果：{detectedOfficeSummary}",
                    $"Multiple Office entries were detected. Avoid auto mode \"{target.DisplayName}\" and select a specific version. Detected: {detectedOfficeSummary}"),
                9 => T(language, "先打开日志目录查看 execution.log 和 SaRA_Files.zip；如果微软场景本身也失败，再考虑做残留清理或切换完整界面版 SaRA。", "Check execution.log and SaRA_Files.zip first; then decide further cleanup."),
                10 => T(language, "请关闭当前程序后，以管理员身份重新运行 M365Tool，再执行卸载。", "Close this app and rerun M365Tool as administrator."),
                66 => T(
                    language,
                    $"你选择的目标版本和设备实际检测结果不一致。建议根据当前检测结果重新选择卸载目标。当前检测结果：{detectedOfficeSummary}",
                    $"Selected target version does not match detected Office. Detected: {detectedOfficeSummary}"),
                67 => T(language, "当前传给 SaRA 的版本参数无效。这通常代表工具映射有误，我们可以继续修参数映射。", "The version parameter passed to SaRA is invalid."),
                68 => T(language, "SaRA 未检测到可卸载 Office。若你确认机器上有残留，下一步建议做残留清理或先检查安装检测结果。", "SaRA found no uninstallable Office. Check detection results or residual cleanup."),
                _ => T(language, "请先打开日志目录查看 execution.log 和 SaRA_Files.zip，我们再根据实际输出继续处理。", "Open log folder and share execution.log + SaRA_Files.zip for next analysis.")
            };
        }

        private static string T(UiLanguage language, string zh, string en)
        {
            return LocalizationService.Localize(language, zh, en);
        }
    }
}
