# 21Vianet Office工具箱

21Vianet Office工具箱是一款面向 21Vianet / 21V Microsoft 365 运维场景的 Windows 桌面工具，用于辅助完成 Microsoft 365 Apps 安装、Office 卸载、更新频道切换、激活痕迹清理、Teams 常见问题处理和 Outlook 诊断等操作。

> 重要说明：本项目仅适用于 21V 用户和 21V 相关 Microsoft 365 环境。工具内置了 21V 账户认证能力，并不适合作为全部 Microsoft 365 租户或环境的通用工具。
>
> 该工具会执行 Microsoft 365 管理和维护相关操作，请先在可控测试环境中验证，再用于生产设备。

## English Summary

21Vianet Office Toolbox is a Windows desktop utility for Microsoft 365 operations in 21Vianet / 21V environments. It includes 21V account verification and is not intended to be a general-purpose tool for all Microsoft 365 tenants.

## 功能模块

- **安装**：配置并安装 Microsoft 365 Apps、Project、Visio，支持版本、位数、频道、语言和排除应用选择。
- **卸载**：常规卸载只提供本机检测到的 C2R 目标，通过 ODT 移除所选版本；无适用目标时禁用开始按钮。深度清理通过微软 Get Help 处理 MSI 或异常安装，支持 Office 2024 等目标。
- **更新频道**：支持当前频道、月度企业频道及旧半年企业频道兼容配置，分别读取微软官方版本列表，支持目标版本和版本回退。
- **清理与修复**：清理激活残留、账户痕迹，并修复代理与网络依赖问题。
- **Teams 工具**：处理 Teams 缓存、登录记录和常见客户端配置异常。
- **Outlook 工具**：执行 Outlook 扫描、诊断、日历检查和日志导出。
- **双语界面**：支持简体中文和英文界面。
- **访问控制**：除首页外的受限功能需要通过 21V 账户验证后使用。

## 运行要求

- Windows 10/11
- 启动工具需要管理员权限，请右键选择“以管理员身份运行”
- 需要访问 Microsoft 相关网络端点，用于验证、安装和诊断
- 21V 账户验证窗口需要 Microsoft Edge WebView2 Runtime

单文件自包含测试版已包含 .NET 运行时；如果系统未安装 WebView2，仍可能需要额外安装 WebView2 Runtime。

## 下载与测试

可以在 GitHub Release 页面下载单文件版：

[下载 21Vianet Office工具箱 v0.1.6](https://github.com/terry-sys/M365Tool/releases/tag/v0.1.6)

下载 `21VianetOfficeToolbox-portable-win-x64-v0.1.6.exe` 后直接运行即可；也可以下载同名 `.zip` 压缩包。

本次更新见 [CHANGELOG.md](CHANGELOG.md)。单文件版包含 .NET 8 运行时；EXE 尚未进行代码签名，Windows 可能显示“未知发布者”。

## 从源码构建

在 Windows 上安装 .NET 8 SDK 后执行：

```powershell
dotnet test M365Tool.sln
dotnet publish src/M365Tool.UI/M365Tool.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -p:IncludeAllContentForSelfExtract=true -p:EnableCompressionInSingleFile=true -p:PublishTrimmed=false -p:DebugType=None -p:DebugSymbols=false -o artifacts/publish
```

发布目录中的 `M365Tool.exe` 可独立分发，首次启动会解压内置组件；系统仍需安装 Microsoft Edge WebView2 Runtime。

## 使用注意事项

- 本工具面向 21V 场景，需要完成 21V 账户验证。
- 本工具不适用于 21V 以外的通用 Microsoft 365 环境。
- 执行会修改本地状态的操作前，建议先关闭 Office、Teams 和 Outlook。
- 安装前请先查看预检查结果。
- 测试失败时请保留导出的日志，便于后续排查。

## 执行与恢复

- 安装默认保留已有 MSI 版 Office、Project 和 Visio；如有兼容性冲突，请先在卸载页明确选择目标处理。频道切换会提示保存文档，默认不强制关闭 Office。
- 频道切换通过 ODT 的 `Updates Channel` 配置提交，再发起手动更新；安装完成后频道才会生效，组织策略具有更高优先级。选择“最新版本”会清除本工具提交的目标版本固定配置。
- 微软自 2606 起统一半年与月度企业频道的更新体验。本工具对已升级的 Microsoft 365 Apps 显示“月度企业频道（原半年配置）”，旧 CDN 及对应 Build 保留兼容，不将两者的 Build 混用。新企业部署可选择月度企业频道，工具原有的当前频道默认选项保持不变。
- 频道规则依据：[微软频道合并说明](https://learn.microsoft.com/en-us/microsoft-365-apps/updates/unified-update-channels)、[ODT 切换频道方式](https://learn.microsoft.com/en-us/microsoft-365-apps/updates/change-update-channels)。
- 同一程序实例中的维护任务互斥。停止等待后，程序会继续跟踪已启动的进程或后台卸载复查，结束前阻止新的维护任务。
- 登录验证在微软认证成功回跳、登录会话建立后立即进入工具；资料页不对用户显示，用户名与头像随后在后台同步。资料同步不影响登录放行。旧版验证缓存需要重新登录一次。
- Outlook 离线扫描必须使用支持 `-OfflineScan` 的诊断工具，不支持时停止操作；扫描结果只从本次运行的输出或日志判定。
- 安装和卸载会先准备工具并验证微软数字签名。卸载默认尝试正常关闭 Office，需要强制关闭时应先保存文档并勾选对应选项。
- 安装和卸载运行时可使用“取消/停止等待”。准备阶段可以取消；维护进程启动后，该按钮停止等待并保留输出，不会强制终止已启动的安装或卸载。请确认后台任务结束后再重试。
- 网络修复会先展示所选更改并备份配置，默认保留 PAC 和现有代理；自动代理检测、WinHTTP 代理替换需单独选择。
- 网络备份保存在当前用户的 `%LOCALAPPDATA%\M365Tool\Sessions\NetworkRepair` 中，可使用“恢复网络备份”选择同一 Windows 用户的备份恢复原值。
- 回归测试运行方式见 [tests/README.md](tests/README.md)。

## 当前状态

项目仍在持续优化中，当前重点是完善 UI 一致性、功能测试体验和 WinForms 层与运维逻辑的边界。
