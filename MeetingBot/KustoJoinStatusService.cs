using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using Azure.Core;

namespace MeetingBot;

/// <summary>Queries Kusto telemetry to determine whether an ISV bot joined a meeting.</summary>
public sealed class KustoJoinStatusService
{
    private const string KustoScope = "https://kusto.kusto.windows.net/.default";
    private const string Query = """
        declare query_parameters(p_threadId:string, p_applicationId:string);
        PT_CMD_scenarions
        | where threadId == p_threadId and customApplicationTokenAppId == p_applicationId
        | take 1
        | project isJoined = true
        """;

    private readonly HttpClient _httpClient;
    private readonly TokenCredential _credential;
    private readonly KustoOptions _options;

    public KustoJoinStatusService(HttpClient httpClient, TokenCredential credential, KustoOptions options)
    {
        _httpClient = httpClient;
        _credential = credential;
        _options = options;
    }

    public async Task<bool> IsJoinedAsync(
        string threadId,
        string customApplicationTokenAppId,
        CancellationToken cancellationToken)
    {
        if (!Uri.TryCreate(_options.ClusterUri, UriKind.Absolute, out Uri? clusterUri))
        {
            throw new InvalidOperationException("Kusto:ClusterUri must be an absolute URI.");
        }

        if (string.IsNullOrWhiteSpace(_options.Database))
        {
            throw new InvalidOperationException("Kusto:Database is required.");
        }

        AccessToken accessToken = await _credential.GetTokenAsync(
            new TokenRequestContext([KustoScope]),
            cancellationToken);

        string properties = JsonSerializer.Serialize(new
        {
            Parameters = new Dictionary<string, string>
            {
                ["p_threadId"] = threadId,
                ["p_applicationId"] = customApplicationTokenAppId,
            },
        });

        string requestBody = JsonSerializer.Serialize(new
        {
            db = _options.Database,
            csl = Query,
            properties,
        });

        Uri queryUri = new Uri(clusterUri, "/v2/rest/query");
        using HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, queryUri)
        {
            Content = new StringContent(requestBody, Encoding.UTF8, "application/json"),
        };
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken.Token);
        request.Headers.Add("x-ms-client-request-id", $"MeetingBot.JoinStatus;{Guid.NewGuid()}");

        using HttpResponseMessage response = await _httpClient.SendAsync(request, cancellationToken);
        string responseBody = await response.Content.ReadAsStringAsync(cancellationToken);
        if (!response.IsSuccessStatusCode)
        {
            throw new HttpRequestException(
                $"Kusto query failed with status {(int)response.StatusCode} ({response.ReasonPhrase}): {responseBody}",
                inner: null,
                response.StatusCode);
        }

        using JsonDocument document = JsonDocument.Parse(responseBody);
        return HasPrimaryResultRow(document.RootElement);
    }

    private static bool HasPrimaryResultRow(JsonElement root)
    {
        if (root.ValueKind == JsonValueKind.Array)
        {
            foreach (JsonElement frame in root.EnumerateArray())
            {
                if (frame.TryGetProperty("TableKind", out JsonElement tableKind)
                    && tableKind.ValueEquals("PrimaryResult")
                    && frame.TryGetProperty("Rows", out JsonElement rows))
                {
                    return rows.GetArrayLength() > 0;
                }
            }
        }

        if (root.ValueKind == JsonValueKind.Object
            && root.TryGetProperty("Tables", out JsonElement tables))
        {
            foreach (JsonElement table in tables.EnumerateArray())
            {
                if (table.TryGetProperty("TableName", out JsonElement tableName)
                    && tableName.ValueEquals("PrimaryResult")
                    && table.TryGetProperty("Rows", out JsonElement rows))
                {
                    return rows.GetArrayLength() > 0;
                }
            }
        }

        throw new InvalidOperationException("Kusto response did not contain a primary result table.");
    }
}
