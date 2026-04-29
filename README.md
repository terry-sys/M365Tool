# 21V M365 助手

21V M365 助手是一款面向 21V Microsoft 365 运维场景的 Windows 桌面工具，用于辅助完成 Microsoft 365 Apps 安装、Office 卸载、更新频道切换、激活痕迹清理、Teams 常见问题处理和 Outlook 诊断等操作。

> 重要说明：本项目仅适用于 21V 用户和 21V 相关 Microsoft 365 环境。工具内置了 21V 账户认证能力，并不适合作为全部 Microsoft 365 租户或环境的通用工具。
>
> 该工具会执行 Microsoft 365 管理和维护相关操作，请先在可控测试环境中验证，再用于生产设备。

## English Summary

21V M365 Assistant is a Windows desktop utility for Microsoft 365 operations in 21V environments. It includes 21V account verification and is not intended to be a general-purpose tool for all Microsoft 365 tenants.

## 功能模块

- **安装**：配置并安装 Microsoft 365 Apps、Project、Visio，支持版本、位数、频道、语言和排除应用选择。
- **卸载**：检测已安装的 Office 产品，并调用微软官方能力执行卸载处理。
- **更新频道**：切换 Office 更新频道，支持按需切换到目标版本或执行版本回退。
- **清理与修复**：清理激活残留、账户痕迹，并修复代理与网络依赖问题。
- **Teams 工具**：处理 Teams 缓存、登录记录和常见客户端配置异常。
- **Outlook 工具**：执行 Outlook 扫描、诊断、日历检查和日志导出。
- **双语界面**：支持简体中文和英文界面。
- **访问控制**：除首页外的受限功能需要通过 21V 账户验证后使用。

## 运行要求

- Windows 10/11，建议 x64 系统
- 使用框架依赖版本时需要 .NET 8 Windows Desktop Runtime
- 执行安装、卸载、清理和修复操作时通常需要管理员权限
- 需要访问 Microsoft 相关网络端点，用于验证、安装和诊断
- 21V 账户验证窗口需要 Microsoft Edge WebView2 Runtime

单文件自包含测试版已包含 .NET 运行时；如果系统未安装 WebView2，仍可能需要额外安装 WebView2 Runtime。

## 下载与测试

可以在 GitHub Release 页面下载单文件版：

[下载 M365Tool v0.1.0](https://github.com/terry-sys/M365Tool/releases/tag/v0.1.0)

下载 `M365Tool-win-x64-single-*.exe` 后直接运行即可。

## 使用注意事项

- 本工具面向 21V 场景，受限功能需要完成 21V 账户验证。
- 本工具不适用于 21V 以外的通用 Microsoft 365 环境。
- 执行会修改本地状态的操作前，建议先关闭 Office、Teams 和 Outlook。
- 安装前请先查看预检查结果。
- 测试失败时请保留导出的日志，便于后续排查。

## 当前状态

项目仍在持续优化中，当前重点是完善 UI 一致性、功能测试体验和 WinForms 层与运维逻辑的边界。
