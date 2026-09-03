using System.Globalization;
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
        | top 1 by EventInfo_Time asc
        | project eventTime = EventInfo_Time
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

    public async Task<DateTimeOffset?> GetJoinEventTimeAsync(
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
        return GetPrimaryResultEventTime(document.RootElement);
    }

    private static DateTimeOffset? GetPrimaryResultEventTime(JsonElement root)
    {
        if (root.ValueKind == JsonValueKind.Array)
        {
            foreach (JsonElement frame in root.EnumerateArray())
            {
                if (frame.TryGetProperty("TableKind", out JsonElement tableKind)
                    && tableKind.ValueEquals("PrimaryResult")
                    && frame.TryGetProperty("Rows", out JsonElement rows))
                {
                    return GetEventTime(frame, rows);
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
                    return GetEventTime(table, rows);
                }
            }
        }

        throw new InvalidOperationException("Kusto response did not contain a primary result table.");
    }

    private static DateTimeOffset? GetEventTime(JsonElement table, JsonElement rows)
    {
        if (rows.GetArrayLength() == 0)
        {
            return null;
        }

        if (!table.TryGetProperty("Columns", out JsonElement columns))
        {
            throw new InvalidOperationException("Kusto primary result table did not contain column metadata.");
        }

        int eventTimeIndex = -1;
        int index = 0;
        foreach (JsonElement column in columns.EnumerateArray())
        {
            if (column.TryGetProperty("ColumnName", out JsonElement columnName)
                && string.Equals(columnName.GetString(), "eventTime", StringComparison.OrdinalIgnoreCase))
            {
                eventTimeIndex = index;
                break;
            }

            index++;
        }

        if (eventTimeIndex < 0)
        {
            throw new InvalidOperationException("Kusto primary result table did not contain the eventTime column.");
        }

        JsonElement row = rows[0];
        if (row.ValueKind != JsonValueKind.Array || row.GetArrayLength() <= eventTimeIndex)
        {
            throw new InvalidOperationException("Kusto primary result row did not contain an eventTime value.");
        }

        string? rawEventTime = row[eventTimeIndex].GetString();
        if (!DateTimeOffset.TryParse(
                rawEventTime,
                CultureInfo.InvariantCulture,
                DateTimeStyles.AssumeUniversal | DateTimeStyles.AdjustToUniversal,
                out DateTimeOffset eventTime))
        {
            throw new InvalidOperationException("Kusto primary result row contained an invalid eventTime value.");
        }

        return eventTime;
    }
}
