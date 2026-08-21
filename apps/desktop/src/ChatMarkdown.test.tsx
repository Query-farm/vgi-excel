import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, it } from "vitest";
import { ChatMarkdown } from "./ChatMarkdown";

describe("ChatMarkdown", () => {
  it("renders Cupola-style GitHub-flavored Markdown", () => {
    const html = renderToStaticMarkup(<ChatMarkdown content={`## Result

- **Clear sky**
- \`temperature\`

| city | °C |
| --- | ---: |
| Paris | 21 |

\`\`\`sql
SELECT 21;
\`\`\`

[Source](https://example.com)`}/>);
    expect(html).toContain("<h3>Result</h3>");
    expect(html).toContain("<strong>Clear sky</strong>");
    expect(html).toContain("<table>");
    expect(html).toContain("language-sql");
    expect(html).toContain("Copy SQL");
    expect(html).toContain('target="_blank"');
  });

  it("does not execute raw HTML or load agent-supplied images", () => {
    const html = renderToStaticMarkup(<ChatMarkdown content={'<script>alert(1)</script>\n\n![remote](https://example.com/pixel.png)'}/>);
    expect(html).not.toContain("<script");
    expect(html).not.toContain("<img");
    expect(html).toContain("[Image: remote]");
  });

  it("renders an incomplete fenced block while text is still streaming", () => {
    const html = renderToStaticMarkup(<ChatMarkdown content={"Here is the query:\n\n```sql\nSELECT 42"} streaming/>);
    expect(html).toContain("language-sql");
    expect(html).toContain("SELECT 42");
    expect(html).not.toContain("```sql");
  });
});
