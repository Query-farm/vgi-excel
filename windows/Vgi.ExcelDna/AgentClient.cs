using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace QueryFarm.Vgi.ExcelDna;

internal sealed class AgentAnswer
{
    public string Text { get; set; } = "";
    public QueryResult? StagedResult { get; set; }
}

internal sealed class AgentClient
{
    private static readonly HttpClient Http = new() { Timeout = TimeSpan.FromMinutes(5) };
    private readonly HaybarnClient _haybarn;

    public AgentClient(HaybarnClient haybarn) => _haybarn = haybarn;

    static AgentClient() => ServicePointManager.SecurityProtocol |= SecurityProtocolType.Tls12;

    public async Task<AgentAnswer> RunAsync(string apiKey, string prompt, string connection)
    {
        if (string.IsNullOrWhiteSpace(apiKey)) throw new ArgumentException("An Anthropic API key is required.");
        if (string.IsNullOrWhiteSpace(prompt)) throw new ArgumentException("An agent prompt is required.");
        var messages = new JArray(new JObject { ["role"] = "user", ["content"] = prompt });
        var definition = ConnectionStore.Resolve(connection);
        QueryResult? inventory = null;
        string? inventoryError = null;
        try
        {
            var catalog = SqlString(definition.Catalog);
            inventory = _haybarn.QueryResult($@"SELECT table_catalog AS catalog, table_schema AS schema, table_name AS name, CASE WHEN table_type='VIEW' THEN 'view' ELSE 'table' END AS kind, '' AS description FROM information_schema.tables WHERE table_catalog={catalog} AND table_schema NOT IN ('information_schema','pg_catalog') UNION ALL SELECT database_name, schema_name, function_name, CASE WHEN function_type IN ('macro','table_macro') THEN 'macro' ELSE function_type END, COALESCE(description, comment, '') FROM duckdb_functions() WHERE database_name={catalog} ORDER BY 1,2,4,3", connection, 20_000);
        }
        catch (Exception error) { inventoryError = error.Message; }
        var systemPrompt = AgentPromptBuilder.Build(definition, inventory, inventoryError);
        var answer = new StringBuilder();
        QueryResult? staged = null;
        for (var round = 0; round < 8; round++)
        {
            var response = await Send(apiKey, systemPrompt, messages);
            var content = (JArray?)response["content"] ?? throw new InvalidOperationException("The agent returned no content.");
            messages.Add(new JObject { ["role"] = "assistant", ["content"] = content.DeepClone() });
            foreach (var text in content.OfType<JObject>().Where(item => (string?)item["type"] == "text")) answer.Append((string?)text["text"]);
            var calls = content.OfType<JObject>().Where(item => (string?)item["type"] == "tool_use").ToArray();
            if (calls.Length == 0) return new AgentAnswer { Text = answer.ToString(), StagedResult = staged };
            var results = new JArray();
            foreach (var call in calls)
            {
                try
                {
                    var tool = RunTool((string?)call["name"] ?? "", (JObject?)call["input"] ?? new JObject(), connection);
                    if (tool.Result is not null) staged = tool.Result;
                    results.Add(new JObject { ["type"] = "tool_result", ["tool_use_id"] = (string?)call["id"], ["content"] = tool.Content });
                }
                catch (Exception error)
                {
                    results.Add(new JObject { ["type"] = "tool_result", ["tool_use_id"] = (string?)call["id"], ["is_error"] = true, ["content"] = error.Message });
                }
            }
            messages.Add(new JObject { ["role"] = "user", ["content"] = results });
        }
        throw new InvalidOperationException("The agent exceeded the maximum of eight tool rounds.");
    }

    private static async Task<JObject> Send(string apiKey, string systemPrompt, JArray messages)
    {
        var body = new JObject
        {
            ["model"] = Environment.GetEnvironmentVariable("VGI_EXCEL_ANTHROPIC_MODEL") ?? "claude-sonnet-4-5",
            ["max_tokens"] = 4096,
            ["system"] = systemPrompt,
            ["tools"] = Tools.DeepClone(),
            ["messages"] = messages.DeepClone()
        };
        using var request = new HttpRequestMessage(HttpMethod.Post, "https://api.anthropic.com/v1/messages");
        request.Headers.Add("x-api-key", apiKey);
        request.Headers.Add("anthropic-version", "2023-06-01");
        request.Content = new StringContent(body.ToString(Formatting.None), Encoding.UTF8, "application/json");
        using var response = await Http.SendAsync(request);
        var value = await response.Content.ReadAsStringAsync();
        if (!response.IsSuccessStatusCode) throw new InvalidOperationException($"Anthropic request failed ({(int)response.StatusCode}): {value}");
        return JObject.Parse(value);
    }

    private (string Content, QueryResult? Result) RunTool(string name, JObject input, string connection)
    {
        switch (name)
        {
            case "run_sql":
            {
                var sql = (string?)input["sql"] ?? "";
                AgentSqlPolicy.AssertReadOnly(sql);
                var result = _haybarn.QueryResult(sql, connection, 10_000);
                return (Summary(result), result);
            }
            case "list_tables": return (Summary(_haybarn.QueryResult(@"SELECT table_catalog, table_schema, table_name, table_type FROM information_schema.tables WHERE table_schema NOT IN ('information_schema', 'pg_catalog') ORDER BY 1, 2, 3", connection, 10_000)), null);
            case "list_functions":
            {
                var functions = _haybarn.QueryResult(@"SELECT database_name, schema_name, function_name, function_type, CAST(to_json(parameters) AS VARCHAR) AS parameters, CAST(to_json(parameter_types) AS VARCHAR) AS parameter_types, return_type, description, comment, CAST(to_json(examples) AS VARCHAR) AS examples, CAST(to_json(tags) AS VARCHAR) AS tags FROM duckdb_functions() WHERE database_name NOT IN ('system', 'temp') ORDER BY 1, 2, 3", connection, 10_000);
                var arguments = _haybarn.QueryResult(@"SELECT catalog_name, schema_name, function_name, function_type, arg_position, field_index, arg_name, arg_type, arg_description, is_named, is_positional, is_const, is_varargs, is_table_input, is_any_type, arg_default, arg_choices, arg_range, arg_pattern FROM vgi_function_arguments() ORDER BY 1, 2, 3, field_index", connection, 10_000);
                return (JsonConvert.SerializeObject(new
                {
                    guidance = "Use positional arguments in arg_position order and pass is_named arguments with name := value. Respect choices, ranges, patterns, and exact types.",
                    functions = SummaryObject(functions),
                    vgi_arguments = SummaryObject(arguments)
                }), null);
            }
            case "describe_table":
            {
                var schema = SqlString((string?)input["schema"] ?? "");
                var table = SqlString((string?)input["table"] ?? "");
                var catalog = string.IsNullOrWhiteSpace((string?)input["catalog"]) ? "" : $" AND table_catalog = {SqlString((string)input["catalog"]!)}";
                return (Summary(_haybarn.QueryResult($"SELECT table_catalog, table_schema, table_name, column_name, data_type, is_nullable FROM information_schema.columns WHERE table_schema = {schema} AND table_name = {table}{catalog} ORDER BY ordinal_position", connection, 10_000)), null);
            }
            default: throw new ArgumentException($"Unknown agent tool: {name}");
        }
    }

    private static string Summary(QueryResult result) => JsonConvert.SerializeObject(new
    {
        columns = result.Columns,
        rows = result.Rows.Take(20),
        row_count = result.RowCount,
        truncated = result.Truncated || result.Rows.Length < result.RowCount
    });

    private static object SummaryObject(QueryResult result) => new
    {
        columns = result.Columns,
        rows = result.Rows.Take(100),
        row_count = result.RowCount,
        truncated = result.Truncated || result.Rows.Length < result.RowCount
    };

    private static string SqlString(string value) => $"'{value.Replace("'", "''")}'";

    private static readonly JArray Tools = JArray.Parse(@"[
      {'name':'run_sql','description':'Run one read-only SQL statement and return a preview.','input_schema':{'type':'object','properties':{'sql':{'type':'string'}},'required':['sql']}},
      {'name':'list_tables','description':'List available catalogs, schemas, tables, and views.','input_schema':{'type':'object','properties':{}}},
      {'name':'list_functions','description':'List every callable plus rich VGI named/positional arguments, constraints, documentation, examples, and tags.','input_schema':{'type':'object','properties':{}}},
      {'name':'describe_table','description':'Describe a table or view and its columns.','input_schema':{'type':'object','properties':{'catalog':{'type':'string'},'schema':{'type':'string'},'table':{'type':'string'}},'required':['schema','table']}}
    ]");
}
