using System;

namespace QueryFarm.Vgi.ExcelDna;

internal static class AgentCredentialStore
{
    private const string DefaultTarget = "QueryFarm/VgiExcel/AnthropicApiKey";

    public static string? Load() => CredentialVault.ReadSecret(Target());

    public static void Save(string apiKey)
    {
        var value = (apiKey ?? "").Trim();
        if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("An Anthropic API key is required.");
        if (value.Length > 2048) throw new ArgumentException("The Anthropic API key is too long.");
        CredentialVault.WriteSecret(Target(), value);
    }

    public static void Delete() => CredentialVault.Delete(Target());

    private static string Target()
    {
        var configured = Environment.GetEnvironmentVariable("VGI_EXCEL_ANTHROPIC_CREDENTIAL_TARGET");
        return string.IsNullOrWhiteSpace(configured) ? DefaultTarget : configured;
    }
}
