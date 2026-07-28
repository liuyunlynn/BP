using Microsoft.AspNetCore.Http.Extensions;

namespace MeetingBot;

/// <summary>
/// Helpers to bridge ASP.NET Core's <see cref="HttpRequest"/>/<see cref="HttpResponse"/>
/// with the <see cref="HttpRequestMessage"/>/<see cref="HttpResponseMessage"/> types the
/// Graph Communications SDK notification pipeline expects.
/// </summary>
public static class HttpTranslation
{
    public static async Task<HttpRequestMessage> ToHttpRequestMessageAsync(this HttpRequest request)
    {
        HttpRequestMessage message = new HttpRequestMessage(new HttpMethod(request.Method), request.GetEncodedUrl());

        MemoryStream bodyStream = new MemoryStream();
        await request.Body.CopyToAsync(bodyStream).ConfigureAwait(false);
        bodyStream.Position = 0;
        message.Content = new StreamContent(bodyStream);

        foreach (KeyValuePair<string, Microsoft.Extensions.Primitives.StringValues> header in request.Headers)
        {
            if (!message.Headers.TryAddWithoutValidation(header.Key, header.Value.ToArray()))
            {
                message.Content.Headers.TryAddWithoutValidation(header.Key, header.Value.ToArray());
            }
        }

        return message;
    }

    public static async Task CopyToAsync(this HttpResponseMessage source, HttpResponse target)
    {
        target.StatusCode = (int)source.StatusCode;

        foreach (KeyValuePair<string, IEnumerable<string>> header in source.Headers)
        {
            target.Headers[header.Key] = header.Value.ToArray();
        }

        if (source.Content is not null)
        {
            foreach (KeyValuePair<string, IEnumerable<string>> header in source.Content.Headers)
            {
                target.Headers[header.Key] = header.Value.ToArray();
            }

            await source.Content.CopyToAsync(target.Body).ConfigureAwait(false);
        }
    }
}
