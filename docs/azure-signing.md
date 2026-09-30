# Azure signing for Cupola

For the configured GitHub-hosted release process, see [GitHub releases](github-releases.md).

The release publisher accepts `-AzureSigningConfigPath` as an alternative to
`-CertificateThumbprint`. Azure signing is prepared but cannot be qualified until
the organization's Public Trust identity and certificate profile are approved.
No test in the default Windows suite sends a signing request.

## While identity approval is pending

1. Record the Artifact Signing account name, subscription ID, tenant ID, and
   region. Tags are optional organizational metadata and do not configure signing.
2. Prepare a dedicated Windows release identity. For GitHub Actions, create an
   Entra app registration and add a federated credential for GitHub's environment
   `cupola-release` in the actual repository. Use issuer
   `https://token.actions.githubusercontent.com`, audience `api://AzureADTokenExchange`,
   and subject the repository's actual immutable subject (including owner/repository IDs); see
   [GitHub releases](github-releases.md).
3. Create that GitHub environment and restrict it to approved release branches
   and reviewers. Configure the dedicated Windows/Excel runner; it needs Azure
   CLI, .NET 8 x64 runtime, and the existing Windows build prerequisites.
4. Install the pinned Microsoft signing module in the build user's context:

   ```powershell
   Install-PackageProvider NuGet -MinimumVersion 2.8.5.201 -Scope CurrentUser -Force
   Install-Module ArtifactSigning -RequiredVersion 0.1.20 -Scope CurrentUser -Force
   ```

   The module installs its signing-tool dependencies on first use, requiring
   access to its package feeds. It signs PE files, MSI files, and PowerShell
   scripts. These tools are build dependencies and are not installed on users'
   desktops. It does not need a downloadable private key or USB token.

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
| `CUPOLA_AZURE_CLIENT_ID` | Secret | Release app's client ID |
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
