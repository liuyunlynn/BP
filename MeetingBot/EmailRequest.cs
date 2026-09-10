using System.Text.Json.Serialization;

namespace MeetingBot
{
    public class EmailRequest
    {
        [JsonPropertyName("to")]
        public string To { get; set; }

        [JsonPropertyName("cc")]
        public string Cc { get; set; }

        [JsonPropertyName("subject")]
        public string Subject { get; set; }

        [JsonPropertyName("body")]
        public string Body { get; set; }

        [JsonPropertyName("isHtml")]
        public bool IsHtml { get; set; }
    }
}
