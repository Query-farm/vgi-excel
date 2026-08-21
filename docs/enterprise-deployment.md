# Enterprise deployment

1. Build the versioned Office bundle with
   `npm run package:office -- --base-url=https://your-origin`, then publish the
   resulting `apps/office/dist` directory.
2. Upload the add-in-only XML manifest in Microsoft 365 Admin Center under
   **Settings → Integrated apps → Upload custom apps** and assign it to a pilot
   group.
3. Publish the signed MSI as an Intune Win32 app for Windows users who require
   legacy XLL functions or native Haybarn execution. It installs no service.
4. Configure every approved VGI service for HTTPS and allow the exact hosted
   add-in origin through CORS. Do not use wildcard CORS.
   Configure the add-in host to return `Cross-Origin-Opener-Policy: same-origin`
   and `Cross-Origin-Embedder-Policy: require-corp`, and serve WASM with
   `application/wasm`.
5. Configure RFC 9728 OAuth protected-resource metadata and public-client PKCE
   on services that require organizational authentication.
6. Deploy the 32-bit or 64-bit packed XLL according to Office bitness and
   configure its trusted location through enterprise Office policy.
7. Ensure the Microsoft Edge WebView2 Evergreen Runtime is deployed and kept
   current. The XLL contains its WebView2 loader and web assets, but not a
   fixed-version browser runtime.
8. Register the Cupola ODBC driver for users who need Power Query/Refresh All.
   If its display name differs from `Cupola for Excel`, deploy
   `CUPOLA_ODBC_DRIVER_NAME` with the registered name. The driver and XLL must
   use the same `CupolaConnection` lookup contract and per-user connection and
   OAuth stores.

The workbook stores connection aliases and HTTPS service URLs only. Microsoft
365 OAuth tokens and the Anthropic BYOK key live in the Office session. The XLL
encrypts its OAuth refresh session with Windows DPAPI for the current user and,
when the user explicitly chooses Save key, stores its Anthropic BYOK key in
Windows Credential Manager. Neither credential is written to workbook formulas
or connection metadata.

Roll out to an IT/power-user group first, validate OAuth and workbook formula
portability in Excel for the web, then expand the assignment group.
