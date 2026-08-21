using System;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace QueryFarm.Vgi.ExcelDna;

internal static class MarkdownRichText
{
    private static readonly Regex Inline = new(@"(\*\*.+?\*\*|`[^`]+`|\[[^\]]+\]\([^)]+\))", RegexOptions.Compiled);

    public static void Set(RichTextBox box, string markdown)
    {
        box.Clear();
        var inCode = false;
        foreach (var raw in (markdown ?? "").Replace("\r\n", "\n").Split('\n'))
        {
            if (raw.TrimStart().StartsWith("```", StringComparison.Ordinal)) { inCode = !inCode; continue; }
            if (inCode) { Append(box, raw + "\n", new Font("Consolas", 9.5f), Color.FromArgb(35, 49, 32)); continue; }
            var line = raw;
            var size = 10f;
            var bold = false;
            if (line.StartsWith("### ")) { line = line.Substring(4); size = 11f; bold = true; }
            else if (line.StartsWith("## ")) { line = line.Substring(3); size = 12f; bold = true; }
            else if (line.StartsWith("# ")) { line = line.Substring(2); size = 14f; bold = true; }
            else if (Regex.IsMatch(line, @"^\s*[-*]\s+")) line = "• " + Regex.Replace(line, @"^\s*[-*]\s+", "");
            AppendInline(box, line, new Font("Segoe UI", size, bold ? FontStyle.Bold : FontStyle.Regular));
            box.AppendText("\n");
        }
        box.SelectionStart = 0;
        box.ScrollToCaret();
    }

    private static void AppendInline(RichTextBox box, string line, Font font)
    {
        var position = 0;
        foreach (Match match in Inline.Matches(line))
        {
            Append(box, line.Substring(position, match.Index - position), font, box.ForeColor);
            var token = match.Value;
            if (token.StartsWith("**")) Append(box, token.Substring(2, token.Length - 4), new Font(font, font.Style | FontStyle.Bold), box.ForeColor);
            else if (token.StartsWith("`")) Append(box, token.Substring(1, token.Length - 2), new Font("Consolas", font.Size), Color.FromArgb(45, 79, 34));
            else
            {
                var close = token.IndexOf("](", StringComparison.Ordinal);
                Append(box, token.Substring(1, close - 1), new Font(font, FontStyle.Underline), Color.FromArgb(55, 101, 37));
                Append(box, " (" + token.Substring(close + 2, token.Length - close - 3) + ")", font, box.ForeColor);
            }
            position = match.Index + match.Length;
        }
        Append(box, line.Substring(position), font, box.ForeColor);
    }

    private static void Append(RichTextBox box, string text, Font font, Color color)
    {
        box.SelectionStart = box.TextLength;
        box.SelectionLength = 0;
        box.SelectionFont = font;
        box.SelectionColor = color;
        box.AppendText(text);
    }
}
