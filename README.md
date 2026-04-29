# 21V M365 Assistant

21V M365 Assistant is a Windows desktop utility for Microsoft 365 operations in 21V environments. It provides a guided workbench for installing Microsoft 365 Apps, uninstalling Office, switching update channels, cleaning activation traces, repairing common Teams issues, and running Outlook diagnostics.

> This tool performs administrative Microsoft 365 maintenance tasks. Test in a controlled environment before using it on production devices.
>
> This project is intended for 21V users only. It includes 21V account verification and is not designed as a general-purpose tool for all Microsoft 365 tenants or environments.

## Features

- **Install**: configure Microsoft 365 Apps, Project, or Visio installation options, including edition, architecture, channel, languages, and excluded apps.
- **Uninstall**: detect installed Office products and run Microsoft-supported uninstall workflows.
- **Update Channel**: switch Office update channels and optionally pin to a target build.
- **Cleanup & Repair**: clean activation remnants, account traces, proxy settings, and network dependencies.
- **Teams Tools**: reset New Teams, clear cache, delete sign-in records, and repair the Teams Meeting Add-in.
- **Outlook Tools**: run Outlook diagnostics, offline scans, calendar checks, and export logs.
- **Bilingual UI**: supports Simplified Chinese and English.
- **Access Gate**: restricted modules require 21V portal verification before use.

## Requirements

- Windows 10/11, x64 recommended
- .NET 8 Windows Desktop Runtime when using framework-dependent builds
- Administrator privileges for install, uninstall, cleanup, and repair operations
- Network access to Microsoft endpoints for validation, installation, and diagnostics
- Microsoft Edge WebView2 Runtime for the portal verification window

Self-contained test builds include the .NET runtime, but WebView2 may still be required by Windows if it is not already installed.

## Download And Test

For local testing, use the generated package under `releases/` if present:

```powershell
.\releases\M365Tool-win-x64-single-*.exe
```

If no release package exists, build one:

```powershell
dotnet publish .\src\M365Tool.UI\M365Tool.csproj `
  -c Release `
  -r win-x64 `
  --self-contained true `
  /p:PublishSingleFile=true `
  /p:IncludeNativeLibrariesForSelfExtract=true `
  /p:EnableCompressionInSingleFile=true
```

The published executable will be created under:

```text
src/M365Tool.UI/bin/Release/net8.0-windows7.0/win-x64/publish/
```

## Build From Source

```powershell
git clone https://github.com/terry-sys/M365Tool.git
cd M365Tool
dotnet restore .\M365Tool.sln
dotnet build .\M365Tool.sln
```

Run from Visual Studio by opening `M365Tool.sln`, or run the project directly:

```powershell
dotnet run --project .\src\M365Tool.UI\M365Tool.csproj
```

## Repository Layout

```text
src/M365Tool.UI/          Windows Forms application
src/M365Tool.UI/Services/ Microsoft 365 operation services
src/M365Tool.UI/Models/   Configuration and result models
```

## Localization

UI text is maintained through `T("中文", "English")` pairs and resource files:

- `src/M365Tool.UI/Resources/UiStrings.resx`
- `src/M365Tool.UI/Resources/UiStrings.zh-CN.resx`

## Safety Notes

- This application is built for 21V scenarios and requires 21V account verification for restricted functions.
- It is not intended for general Microsoft 365 environments outside 21V.
- Run as administrator when performing installation, uninstall, cleanup, or repair actions.
- Close Office, Teams, and Outlook before running operations that modify their local state.
- Review pre-check output before installing Microsoft 365 Apps.
- Keep exported logs when testing failures; they are useful for diagnosis.
- Do not commit local build outputs from `bin/`, `obj/`, `artifacts/`, or `releases/`.

## Status

This repository is actively being refined. Current work focuses on UI consistency, test packaging, and separating operation logic from the WinForms layer.
