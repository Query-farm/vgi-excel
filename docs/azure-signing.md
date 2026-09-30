# Azure signing for Cupola

The release publisher accepts `-AzureSigningConfigPath` as an alternative to
`-CertificateThumbprint`. Workstation signing was verified on Europa on
2026-09-30 using the approved Public Trust profile `cupola-production`.
No test in the default Windows suite sends a signing request.

## GitHub Actions signing

The configured release workflow uses the user-assigned managed identity
`cupola-github-releases` and GitHub OIDC. It does not use a client secret or the
workstation user's Azure session. The identity's only role is Artifact Signing
Certificate Profile Signer on `cupola-production`. See [GitHub releases](github-releases.md)
for workflow modes, environment configuration, and the exact immutable subject.

Run **Actions → Cupola for Excel Windows release → Run workflow**, selecting
**signing-test** to verify authentication or **signed-draft** to build a release.
The latter validates the unsigned candidate before signing and creates a draft
for real Excel and installation qualification. It never publishes automatically.
The complete hosted build and signing path passed on 2026-09-30 and created
signed draft `v0.5.0-20260930.0`; all 16 Authenticode payloads passed publisher
and timestamp verification.

## Workstation prerequisites

Europa has Azure CLI, .NET 8 x64, and the pinned Microsoft signing module:

```powershell
Install-PackageProvider NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force
Install-Module ArtifactSigning -RequiredVersion 0.1.20 -Scope CurrentUser -Force
```

The module downloads its signing-tool dependencies on first use. These are build
tools, not customer prerequisites; no downloadable private key or USB token is
needed.

## Once approved

Create a **Public Trust** certificate profile, such as `cupola-production`.
Assign **Artifact Signing Certificate Profile Signer** on that profile to the
release app's service principal. For workstation signing, assign the same role
to the user who will run `az login`. Copy the full certificate subject from the
approved profile; the publisher compares it exactly with the resulting signature
subject. Do not substitute a display name or guess the legal entity fields.

Copy `windows/signing/azure.example.json` to an ignored location such as
`artifacts/azure-signing.json`. Replace the account and subject placeholders and
set the endpoint to the account's region. The East US endpoint in the example
is only an example. This file contains no credentials.

For a first workstation release:

```powershell
az login --tenant YOUR_TENANT_ID
az account set --subscription YOUR_SUBSCRIPTION_ID
.\windows\publish.ps1 `
  -HaybarnPath .\artifacts\native-inputs\haybarn.exe `
  -VgiExtensionPath .\artifacts\native-inputs\vgi.duckdb_extension `
  -OdbcDriverPath .\artifacts\native-inputs\haybarn_odbc.dll `
  -BuildMsi -Production -AzureSigningConfigPath .\artifacts\azure-signing.json
```

For GitHub Actions, configure these in the `cupola-release` environment:

| Setting | Type | Value |
|---|---|---|
| `CUPOLA_AZURE_CLIENT_ID` | Secret | Release managed identity's client ID |
| `CUPOLA_AZURE_TENANT_ID` | Secret | Entra tenant ID |
| `CUPOLA_AZURE_SUBSCRIPTION_ID` | Secret | Signing subscription ID |
| `CUPOLA_AZURE_ENDPOINT` | Variable | Exact regional HTTPS signing endpoint |
| `CUPOLA_AZURE_ACCOUNT` | Variable | Artifact Signing account name |
| `CUPOLA_AZURE_PROFILE` | Variable | Approved Public Trust profile name |
| `CUPOLA_AZURE_SUBJECT` | Variable | Exact approved certificate subject |

The IDs are identifiers, not access credentials; the login action receives them
from protected environment secrets. No client secret is required. Run the Windows
release workflow with mode **signed-draft**. It validates
the unsigned candidate first, logs in with GitHub OIDC, then signs the production
package through the resulting Azure CLI identity. Other credential fallbacks,
including browser prompts, are disabled during signing.

Every signed file must have a valid Authenticode signature, a Microsoft RFC3161
timestamp, and the configured publisher subject. Certificates rotate, so Azure
verification does not pin a short-lived certificate thumbprint. The Haybarn-signed
`vgi.duckdb_extension` is never modified by Authenticode signing.

The first real signed release must still pass signature/trust checks on a clean
Windows desktop and the installation lifecycle/OAuth gates in
[production readiness](windows-production-readiness.md). Offline tests do not prove
Azure authorization or successful signing.

Sources: [Microsoft signing integrations](https://learn.microsoft.com/en-us/azure/artifact-signing/how-to-signing-integrations),
[Microsoft signing module](https://www.powershellgallery.com/packages/ArtifactSigning/0.1.20),
[Azure login with OIDC](https://github.com/Azure/login#login-with-openid-connect-oidc-recommended).

## Preparation verified on Europa

Build 20260922.6 passes the Windows suite, including the offline signing adapter
checks; the workflow passes `actionlint`. Microsoft ArtifactSigning 0.1.20 and
its signing-tool dependencies are installed on Europa, and .NET 8 x64 is present.
Azure CLI has also been installed. Open a new terminal after CLI installation so
`az` is on PATH. No Azure sign-in, signing request, role assignment, or cloud
resource change was performed during this preparation. Account/region and the
approved profile subject still need to be supplied.

## Live workstation verification — 2026-09-30

The `queryfarmsigning` account in East US and its `cupola-production` Public Trust
profile are active. The workstation user has the Artifact Signing Certificate
Profile Signer role. A disposable PowerShell script was signed through Cupola's
existing signing adapter with ArtifactSigning 0.1.20. Windows reported a valid
Authenticode signature, the exact publisher subject below, and a Microsoft
RFC3161 timestamp:

```text
CN=Query Farm LLC, O=Query Farm LLC, L=Glen Allen, S=Virginia, C=US
```

Europa's configuration is saved outside source control at
`C:\Users\rusty\vgi-excel-cupola-test\artifacts\azure-signing.json`, using
`https://eus.codesigning.azure.net/`. No credentials are stored in this file.

The sign-in completed in the interactive Windows session. Signing from SSH
could list the account but could not acquire its Azure CLI credentials; running
the same test in the signed-in interactive session succeeded. Use that Windows
session for workstation releases. GitHub OIDC signing uses its own release environment and managed identity;
it does not inherit this workstation session.

This was a disposable signing test, not a signed release or installation. The
clean-machine trust and installation lifecycle gates still apply to the first
production package.
