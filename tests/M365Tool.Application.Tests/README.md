# M365Tool.Application.Tests

Regression tests for the maintenance services. Run `dotnet test M365Tool.sln` from the repository root on Windows with the .NET 8 SDK.

The suite covers full Office build versions, mandatory Teams repair arguments, inventory filtering, ODT removal scope, signature rejection, uninstall post-checks, process timeout/cancellation, and proxy failure handling with backup restoration.

Functional safety regressions also cover preserving MSI products during installation, non-forced Office updates, current and legacy version-table columns, failed portal navigation and old verification caches, shared maintenance operation lifetimes, strict Outlook result parsing, current-session logs, explicit offline scanning, and execution of a harmless command script under Unicode paths. Version parsing uses a checked-in extract from the Microsoft update-history page, without network access.

Maintenance commands use fake runners. Process tests launch only harmless, short-lived test children. Backup tests create and remove isolated `HKCU\Software\M365Tool.Tests\<GUID>` keys and temporary files, leaving actual proxy/Internet Settings untouched.

Credential sign-in tests verify immediate handoff without a rendered profile, rejection of incomplete/failed authentication, the concealed profile-page policy, trusted account-capture origins and authenticated-session cookie classification. Profile display-name and avatar availability are independent of successful handoff.

The separate portal browser regression suite runs with Node.js and Playwright: `node tests/PortalLoginRegressionTests.cjs` from the repository root. It evaluates the actual C# scripts in an isolated headless Chromium browser, serving all requests from local fixtures. It covers account capture through a fast authentication redirect with a server-established test session and an empty profile page, confirms that passwords are never captured, and checks Chinese/English profile cards, the modern account menu, missing/hidden markers and incorrect origins for the optional identity fallback. It never connects to the live account portal. If needed, set `NODE_PATH` to your existing Playwright module directory and `M365_BROWSER_EXECUTABLE` to your Chromium executable path.
