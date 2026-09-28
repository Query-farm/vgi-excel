# Enterprise deployment

## Windows desktop package

The 0.5.0 candidate targets **x64 Windows with x64 Microsoft 365 Click-to-Run
Excel**. Production qualification is tracked in
[windows-production-readiness.md](windows-production-readiness.md). It is not
approved for broad rollout until those gates pass.

The machine MSI installs Cupola, its Excel startup loader, native Haybarn, the
signed VGI extension, WebView2 loader, complete desktop web assets, and the
Cupola ODBC driver under `%ProgramFiles%\Query Farm\Cupola for Excel`. The COM
loader is registered machine-wide and activates the XLL when each user opens
Excel. There is no service or local network companion. The MSI does not write
per-user Excel startup entries under the SYSTEM account.

Deploy .NET Framework 4.8+, x64 Microsoft 365 Excel, and the **machine-wide**
Microsoft Edge WebView2 Evergreen Runtime before Cupola. Deploy the release
publisher trust and applicable Office add-in policies through your normal IT
policy system. Do not disable Office security globally. These are deployment
prerequisites, not runtime dependencies on Node, Python, Visual Studio, or the
.NET SDK.

`windows/publish.ps1 -BuildMsi` creates `artifacts/deployment` containing the MSI,
deployment scripts, and `deployment.json`. Production releases require
`-Production` with either `-CertificateThumbprint <thumbprint>` or
`-AzureSigningConfigPath <configuration.json>`, plus verified native provenance.
See [Azure signing setup](azure-signing.md).
Signing uses the certificate store/signing provider on the release machine;
private keys are never packaged. PowerShell scripts, the MSI, and PE executable/DLL payloads receive Authenticode signatures and verified timestamps. The VGI extension retains its existing Haybarn signature byte-for-byte; applying Authenticode to that file would invalidate its engine signature.

For Intune, wrap that deployment directory with Microsoft's Win32 Content Prep
Tool. Configure installation behavior **System**:

```powershell
powershell.exe -NoProfile -File Deploy-Cupola.ps1 -Action Install
powershell.exe -NoProfile -File Deploy-Cupola.ps1 -Action Uninstall
powershell.exe -NoProfile -File Deploy-Cupola.ps1 -Action Repair
```

Use `Detect-Cupola.ps1 -MinimumVersion 0.5.0` as the detection script, running as
64-bit PowerShell. It returns success and text only when the required version,
files, loader registration, and ODBC registration are present. Configure exit
code `0` as success, `3010` as soft reboot, and `1618` as retry. The scripts never
close Excel or reboot a desktop. The MSI independently rejects changes while
Excel is running. `-AllowUnsigned` is for isolated qualification only and must
not appear in production deployment commands.

Direct MSI installation also supports standard Windows Installer options:

```bat
msiexec /i CupolaForExcel.msi /qn /norestart /L*v "%TEMP%\Cupola-install.log"
```

The wrapper provides a useful retry code for Intune when Excel is open. A race
where Excel starts after the wrapper check is caught by the MSI and reported as
an installation failure; reschedule after Excel closes. Native ODBC files may
also be in use outside Excel; Windows Installer can report a reboot requirement.

See Microsoft's [Win32 app deployment documentation](https://learn.microsoft.com/en-us/intune/app-management/deployment/add-win32)
and [Windows Installer command reference](https://learn.microsoft.com/en-us/windows-server/administration/windows-commands/msiexec).

## Connection defaults and user data

IT can provision a JSON array of non-secret defaults with the following shape:

```json
[
  {
    "Name": "Company Analytics",
    "Catalog": "analytics",
    "Location": "https://vgi.example.com",
    "Authentication": "oauth",
    "AttachOptions": {}
  }
]
```

After installing Cupola, run `Set-CupolaDefaults.ps1 -Path <approved-file.json>`
as administrator using 64-bit PowerShell. It validates the file and stores it
as `connection-defaults.json` beside the protected program files. Never include
passwords, keys, tokens, or authorization headers. On the next Excel startup,
Cupola validates the defaults again and adds missing connection names to that
user's existing Cupola store. Matching names are case-insensitive; **existing
user settings and the active connection are preserved**. This is additive
provisioning, not policy enforcement. To change an existing connection's
endpoint/options, use a separately reviewed migration rather than silently
redirecting workbooks. A removed default can be re-added at startup while it
remains in the administrator's defaults file.

OAuth sign-in happens in the user's session. Refresh sessions remain encrypted
with Windows DPAPI for that user; saved Anthropic keys remain in Windows
Credential Manager. The MSI never creates, reads, replaces, or deletes these
credentials. Power Query stores SQL and a friendly connection name; its driver
resolves the same per-user connection definition and performs the HTTPS VGI
`ATTACH`. Excel may require first-use **Default or Custom** ODBC permission with
no additional credentials. Automatic installation does not bypass that consent
or organizational OAuth sign-in.

User connection writes retain a previous `.bak` file. If the registry becomes
corrupt, Cupola reports an error instead of writing an empty registry. Recovery
must be performed with Excel closed, from a reviewed backup. Backups and local
connection metadata must not be committed or sent to remote telemetry.

## Updates, removal, and migration

Every production MSI release must increase the three-part `package.json`
version, in addition to `cupolaBuild`. The package rejects downgrades; failed
major upgrades use the Windows Installer rollback transaction. Repair restores
package-owned files and registrations. Uninstall removes the machine loader and
ODBC registrations and package-owned files, preserving user state and externally
provisioned defaults. Existing Haybarn drivers and DSNs remain separate.

The developer updater remains `artifacts\xll\Update Cupola for Excel.cmd` for
unmanaged developer installations. It refuses to replace a machine MSI install.
When migrating users from a developer XLL installation, deploy
`Remove-DeveloperRegistration.ps1` once **in each user's context**, then close
Excel and install the machine MSI. That script removes only Cupola startup
registration; it preserves other add-ins and all connections and credentials.
Do not combine the machine loader with an old per-user Cupola startup entry.

Start with a pilot group and validate standard-user startup, protected-catalog
Power Query refresh, upgrade, repair, and uninstall before broad assignment.

## Microsoft 365 Office add-in

The Office task pane remains a separate deployment: build
`npm run package:office -- --base-url=https://your-origin`, publish the bundle,
and upload the XML manifest through Microsoft 365 Admin Center's Integrated
apps. The HTTPS host must serve WASM correctly, provide the documented COOP/COEP
headers, and use exact allowed CORS origins. Office OAuth material stays in the
Office session. See [development.md](development.md).

## Deploying a multi-catalog profile

Connection defaults can include a profile alongside its members. For example:

```json
[
  {"Name":"Weather","Catalog":"open_meteo","Location":"https://vgi-open-meteo.rusty-bb6.workers.dev","Authentication":"anonymous","AttachOptions":{}},
  {"Name":"Earthquakes","Catalog":"earthquakes","Location":"https://vgi-earthquakes.rusty-bb6.workers.dev","Authentication":"anonymous","AttachOptions":{}},
  {"Name":"Research","Members":["Weather","Earthquakes"],"Location":"","Authentication":"anonymous","AttachOptions":{}}
]
```

The member order selects the default catalog. Import validates all references
before writing defaults. Existing user definitions remain authoritative; their
catalog aliases must still be distinct within each imported profile. Authentication
is managed separately for each member and is never included in this file.

Validate a defaults file without writing machine configuration:

```powershell
.\Set-CupolaDefaults.ps1 -Path .\connection-defaults.json -ValidateOnly
```

Omit `-ValidateOnly` to provision the validated file from an administrator session.
