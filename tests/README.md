# Tests

Run the automated regression tests from the repository root:

```powershell
dotnet test M365Tool.sln
```

`M365Tool.Application.Tests` checks Office version arguments, ODT removal scope, signature rejection, uninstall post-checks, inventory filtering, Teams repair arguments, cancellation/timeouts, and network backup/restore.

Maintenance commands are mocked. Execution tests start short-lived `cmd.exe`/PowerShell test children and clean them up. Backup tests use disposable GUID keys under `HKCU\Software\M365Tool.Tests` and temporary files; they do not change actual Internet Settings or WinHTTP settings. No tests install or uninstall Office.
