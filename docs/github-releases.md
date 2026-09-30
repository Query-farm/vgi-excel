# Cupola for Excel releases in GitHub Actions

Open **Actions → Cupola for Excel Windows release → Run workflow**, select the
release branch, and choose a mode:

- **signed-draft** runs web checks, builds pinned native inputs, validates the
  Windows package, signs it in Azure, and creates a draft GitHub release.
- **candidate** runs the same checks and uploads an unsigned candidate for testing.
- **signing-test** signs a disposable PowerShell file to verify GitHub OIDC,
  signing permissions, the publisher subject, and the Microsoft timestamp.

Release builds run on GitHub-hosted `windows-2022` runners; web checks run on
Ubuntu. Europa and a personal Azure login are not required. Native builds are
cached by the source lock, Cupola ODBC patch, and build script; a cold build takes
longer. Native provenance is verified again before production signing.

## Drafts and qualification

Signed drafts include the MSI, a deployment zip, a developer-updater zip,
SHA-256 checksums, the XLL release manifest, and native provenance. Draft names
include both `version` and `cupolaBuild` from `package.json`. Increment
`cupolaBuild` before making an installable update. The workflow refuses to
replace an existing release with the same tag.

GitHub-hosted Windows does not include Excel. CI runs `tests/run-windows.ps1`
with `-SkipExcel`; passing CI is not a claim that real Excel tests passed.
Before publishing a draft, qualify those exact signed artifacts on Windows with
Excel: formulas/spills, Power Query refresh, interactive Credential Manager,
install/upgrade/uninstall, and active-install diagnostics. Record the results
in the draft release. Publish the draft in GitHub only when those checks pass;
do not rebuild the qualified files. See [Windows readiness](windows-production-readiness.md)
and [test instructions](../tests/README.md).

## Azure identity

The user-assigned managed identity `cupola-github-releases` lives in
`rg-query-farm-signing` in the signing subscription. Its only assigned role is
**Artifact Signing Certificate Profile Signer**, scoped to
`queryfarmsigning/certificateProfiles/cupola-production`. It has no client secret
and needs no interactive login. Azure Login exchanges GitHub's short-lived OIDC
assertion for an Azure CLI session on the runner.

The federation uses issuer `https://token.actions.githubusercontent.com`, audience
`api://AzureADTokenExchange`, and this repository's immutable subject:

```text
repo:Query-farm@183420031/vgi-excel@1341795270:environment:cupola-release
```

The `cupola-release` GitHub environment stores the three Azure identifier secrets
and four signing variables listed in [Azure signing](azure-signing.md). Its branch
policy permits `main` and the setup branch `codex/github-releases`. Remove the
setup-branch policy after merging and qualification. Signing jobs alone receive
`id-token: write`; the build jobs have read-only repository access. The draft
release job also receives `contents: write`, and always uses `--draft`. A separate
input-download job receives `contents: write` because GitHub requires it to read
draft release assets; that job runs no repository code and has no Azure access.

## Approved VGI input

The existing engine-signed `vgi.duckdb_extension` is stored as an asset of the
**draft** release `cupola-native-inputs-v1`. Keep that infrastructure release as a
draft. The workflow downloads it using the repository-scoped GitHub token and
checks its SHA-256 against `windows/native-inputs.lock.json` before building.
The extension is never re-signed with Authenticode. Review changes to the native
lock and the approved input together; never bypass the checksum check.

Sources: [GitHub OIDC subjects](https://docs.github.com/en/actions/reference/security/oidc),
[Azure OIDC authentication](https://learn.microsoft.com/en-us/azure/developer/github/connect-from-azure-openid-connect).
