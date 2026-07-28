using System.Text.Json;
using Microsoft.Graph.Contracts;
using Microsoft.Graph.Models;

namespace MeetingBot;

/// <summary>
/// Parses a Teams meeting join URL (the joinWebUrl returned by Graph) into the
/// <see cref="ChatInfo"/> and <see cref="MeetingInfo"/> objects the calling SDK
/// needs to join the meeting.
/// </summary>
public static class JoinUrlParser
{
    private const string MeetupMarker = "/meetup-join/";

    public static (ChatInfo ChatInfo, MeetingInfo MeetingInfo) Parse(string joinWebUrl)
    {
        if (string.IsNullOrWhiteSpace(joinWebUrl))
        {
            throw new ArgumentException("Join URL must not be empty.", nameof(joinWebUrl));
        }

        Uri uri = new Uri(joinWebUrl);

        int markerIndex = uri.AbsolutePath.IndexOf(MeetupMarker, StringComparison.OrdinalIgnoreCase);
        if (markerIndex < 0)
        {
            throw new FormatException($"Join URL '{joinWebUrl}' is not a recognized Teams meetup-join URL.");
        }

        string afterMarker = uri.AbsolutePath.Substring(markerIndex + MeetupMarker.Length);
        string[] pathParts = afterMarker.Split('/', StringSplitOptions.RemoveEmptyEntries);
        string threadId = Uri.UnescapeDataString(pathParts[0]);
        string messageId = pathParts.Length > 1 ? Uri.UnescapeDataString(pathParts[1]) : "0";

        (string? tenantId, string? organizerId) = ParseContext(uri.Query);

        ChatInfo chatInfo = new ChatInfo
        {
            ThreadId = threadId,
            MessageId = messageId,
        };

        Identity organizer = new Identity
        {
            Id = organizerId,
            DisplayName = "Meeting Organizer",
        };

        if (!string.IsNullOrEmpty(tenantId))
        {
            organizer.SetTenantId(tenantId);
        }

        OrganizerMeetingInfo meetingInfo = new OrganizerMeetingInfo
        {
            Organizer = new IdentitySet
            {
                User = organizer,
            },
        };

        return (chatInfo, meetingInfo);
    }

    private static (string? TenantId, string? OrganizerId) ParseContext(string query)
    {
        if (string.IsNullOrEmpty(query))
        {
            return (null, null);
        }

        string trimmed = query.TrimStart('?');
        foreach (string pair in trimmed.Split('&', StringSplitOptions.RemoveEmptyEntries))
        {
            int eq = pair.IndexOf('=');
            if (eq < 0)
            {
                continue;
            }

            string key = pair.Substring(0, eq);
            if (!string.Equals(key, "context", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            string contextJson = Uri.UnescapeDataString(pair.Substring(eq + 1));
            using JsonDocument document = JsonDocument.Parse(contextJson);
            string? tid = document.RootElement.TryGetProperty("Tid", out JsonElement tidElement) ? tidElement.GetString() : null;
            string? oid = document.RootElement.TryGetProperty("Oid", out JsonElement oidElement) ? oidElement.GetString() : null;
            return (tid, oid);
        }

        return (null, null);
    }
}
