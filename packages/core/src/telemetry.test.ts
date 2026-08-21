import { describe, expect, it } from "vitest";
import { classifyTelemetryError, isTelemetrySensitiveKey, sanitizeTelemetryText, sanitizeTelemetryValue } from "./telemetry.js";

describe("telemetry privacy", () => {
  it("removes SQL rather than attempting a partial scrub", () => {
    const sql = `Binder Error while running SELECT * FROM "Payroll FY26" WHERE employee = 'Ada'`;
    expect(sanitizeTelemetryText(sql)).toBe("[query details redacted]");
  });

  it("removes endpoints, credentials, local paths, and quoted identifiers", () => {
    const source = `Bearer abc.def.ghi at https://weather.example.test/v1?q=secret from C:\\Users\\rusty\\Books\\Payroll.xlsx sheet “Executive Pay”`;
    const safe = sanitizeTelemetryText(source);
    expect(safe).not.toContain("abc.def.ghi");
    expect(safe).not.toContain("weather.example.test");
    expect(safe).not.toContain("rusty");
    expect(safe).not.toContain("Executive Pay");
  });

  it("drops sensitive structured fields and bounds collections", () => {
    const source = {
      operation: "query.execute",
      sql: "SELECT salary FROM payroll",
      workbookName: "Board.xlsx",
      endpoint: "https://private.example.test",
      detail: { message: "HTTP 500", api_key: "sk-ant-secret-secret" },
      items: Array.from({ length: 30 }, (_, index) => index),
    };
    const safe = sanitizeTelemetryValue(source) as Record<string, unknown>;
    expect(safe.operation).toBe("query.execute");
    expect(safe.sql).toBe("[redacted]");
    expect(safe.workbookName).toBe("[redacted]");
    expect(safe.endpoint).toBe("[redacted]");
    expect((safe.detail as Record<string, unknown>).api_key).toBe("[redacted]");
    expect((safe.items as unknown[])).toHaveLength(20);
  });

  it("recognizes workbook and authentication fields regardless of casing", () => {
    expect(isTelemetrySensitiveKey("refresh_token")).toBe(true);
    expect(isTelemetrySensitiveKey("workbookName")).toBe(true);
    expect(isTelemetrySensitiveKey("authorization")).toBe(true);
  });

  it("classifies failures without returning customer-controlled text", () => {
    expect(classifyTelemetryError(new Error("Binder Error: column Employee_SSN missing from Payroll"))).toBe("SQL binder error");
    expect(classifyTelemetryError("Worksheet Executive Bonuses already exists")).toBe("Excel operation failure");
    expect(classifyTelemetryError("Acme North confidential failure marker")).toBe("Unexpected application error");
  });
});
