using System.Web;

namespace QueryFarm.Vgi.ExcelDna;

internal static class OAuthCallbackPage
{
    private const string Style = @"
<style>
  * { margin: 0; padding: 0; box-sizing: border-box; }
  body { font-family: Inter, system-ui, -apple-system, 'Segoe UI', sans-serif; min-height: 100vh; display: flex; flex-direction: column; align-items: center; justify-content: center; padding: 2rem; }
  @media (prefers-color-scheme: light) {
    body { background: #faf8f0; color: #2c2c1e; }
    .card { background: #fff; border: 1px solid #f0ece0; }
    .icon-circle-success { background: #e8f5e0; color: #4a7c23; }
    .icon-circle-error { background: #fde8e8; color: #c53030; }
    .subtitle, .footer { color: #6b6b5a; }
    .resource { color: #2c2c1e; }
    .footer a { color: #4a7c23; }
    .brand-copy small { color: #6b6b5a; }
  }
  @media (prefers-color-scheme: dark) {
    body { background: #1a1a0e; color: #f5f0e0; }
    .card { background: #252518; border: 1px solid #3a3a28; }
    .icon-circle-success { background: #2d501640; color: #6ba034; }
    .icon-circle-error { background: #c5303020; color: #fc8181; }
    .subtitle, .footer { color: #b8b0a0; }
    .resource { color: #f5f0e0; }
    .footer a { color: #6ba034; }
    .brand-copy small { color: #b8b0a0; }
  }
  .card { max-width: 560px; width: 100%; border-radius: 12px; padding: 2.5rem; box-shadow: 0 4px 24px rgba(0,0,0,.06); display: flex; align-items: center; gap: 2rem; }
  .brand { display: flex; flex-direction: column; align-items: center; gap: .6rem; flex: 0 0 132px; text-align: center; }
  .brand svg { width: 92px; height: 92px; }
  .brand-copy strong { display: block; font-size: 1.05rem; }
  .brand-copy small { display: block; margin-top: .15rem; font-size: .72rem; line-height: 1.3; }
  .content { min-width: 0; text-align: left; }
  .status-row { display: flex; align-items: center; gap: .625rem; margin-bottom: .5rem; }
  .icon-circle { width: 36px; height: 36px; border-radius: 50%; display: flex; align-items: center; justify-content: center; flex-shrink: 0; }
  .icon-circle svg { width: 20px; height: 20px; }
  h1 { font-size: 1.375rem; font-weight: 600; }
  .subtitle { font-size: .938rem; line-height: 1.5; }
  .resource { font-weight: 600; overflow-wrap: anywhere; }
  .resource-url { margin-top: .25rem; font: .8125rem 'Cascadia Mono', 'JetBrains Mono', monospace; opacity: .7; overflow-wrap: anywhere; }
  .footer { margin-top: 2rem; font-size: .8125rem; text-align: center; }
  .footer a { text-decoration: none; }
  @media (max-width: 560px) {
    body { padding: 1.25rem; }
    .card { flex-direction: column; align-items: flex-start; gap: 1.25rem; padding: 1.75rem; }
    .brand { flex-basis: auto; flex-direction: row; text-align: left; }
    .brand svg { width: 64px; height: 64px; }
  }
</style>";

    private const string Mark = @"<svg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 64 64' role='img' aria-label='Cupola'><polygon points='32,14 8,32 32,42' fill='#d9a441'/><polygon points='32,14 56,32 32,42' fill='#a9762e'/><polygon points='17,35 32,42 32,59 17,52' fill='#f0c877'/><polygon points='47,35 32,42 32,59 47,52' fill='#7a5230'/><path d='M32 14V10.8' stroke='currentColor' stroke-width='2.2' stroke-linecap='round'/><circle cx='32' cy='8.2' r='2.6' fill='currentColor'/></svg>";

    public static string Success(string resourceDisplay, string resourceUrl) => Page(
        "Authentication Successful", "success",
        "<p class='subtitle'>Connected to <span class='resource'>" + Html(resourceDisplay) + "</span></p>" +
        "<p class='subtitle resource-url' title='" + Html(resourceUrl) + "'>" + Html(resourceUrl) + "</p>" +
        "<p class='subtitle' style='margin-top:.75rem'>You can close this window and return to Excel.</p>");

    public static string Error(string message, string resourceDisplay) => Page(
        "Authentication Failed", "error",
        "<p class='subtitle'>" + Html(message) + "</p>" +
        "<p class='subtitle' style='margin-top:.5rem'>Resource: <span class='resource'>" + Html(resourceDisplay) + "</span></p>");

    private static string Page(string title, string kind, string body)
    {
        var icon = kind == "success"
            ? "<path stroke-linecap='round' stroke-linejoin='round' d='M5 13l4 4L19 7'/>"
            : "<path stroke-linecap='round' stroke-linejoin='round' d='M6 18L18 6M6 6l12 12'/>";
        return "<!DOCTYPE html><html lang='en'><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'><title>" +
            Html(title) + "</title>" + Style + "</head><body><div class='card'><div class='brand'>" + Mark +
            "<div class='brand-copy'><strong>Cupola</strong><small>for Excel</small></div></div><div class='content'><div class='status-row'><div class='icon-circle icon-circle-" + kind +
            "'><svg fill='none' stroke='currentColor' stroke-width='2.5' viewBox='0 0 24 24'>" + icon + "</svg></div><h1>" + Html(title) +
            "</h1></div>" + body + "</div></div><div class='footer'>&copy; 2026 &#x1F69C; <a href='https://query.farm'>Query Farm LLC</a></div></body></html>";
    }

    private static string Html(string value) => HttpUtility.HtmlEncode(value ?? "");
}
