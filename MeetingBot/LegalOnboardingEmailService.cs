using System.Net.Http.Headers;
using System.Net.Mail;
using System.Text;
using System.Text.Encodings.Web;
using System.Text.Json;
using Azure.Core;
using Azure.Identity;

namespace MeetingBot;

/// <summary>Sends Teams Bot Identification Program legal onboarding email through Microsoft Graph.</summary>
public sealed class LegalOnboardingEmailService
{
    private const string Subject = "Teams Bot Identification Program Onboarding";

    private readonly HttpClient _httpClient;
    private readonly ILogger<LegalOnboardingEmailService> _logger;
    private readonly string _apiKey;
    private readonly string _requestUri;

    public LegalOnboardingEmailService(
        LegalOnboardingEmailOptions options,
        IConfiguration configuration,
        IHttpClientFactory httpClientFactory,
        ILogger<LegalOnboardingEmailService> logger)
    {
        _httpClient = httpClientFactory.CreateClient(nameof(LegalOnboardingEmailService));
        _logger = logger;

        string? apiKey =  configuration["LegalOnboardingEmail:ApiKey"]
            ?? options.ApiKey;

        _apiKey = string.IsNullOrWhiteSpace(apiKey)
            ? throw new InvalidOperationException(
                "LegalOnboardingEmail:ApiKey must be configured in appsettings.json or an App Service environment variable.")
            : apiKey.Trim();

        string? requestUri = configuration["LegalOnboardingEmail:RequestUri"]
            ?? options.RequestUri;

        _requestUri = string.IsNullOrWhiteSpace(requestUri)
            ? throw new InvalidOperationException(
                "LegalOnboardingEmail:RequestUri must be configured in appsettings.json or an App Service environment variable.")
            : requestUri.Trim();
    }

    public async Task SendAsync(LegalOnboardingEmailRequest email, CancellationToken cancellationToken)
    {
        Validate(email);

        EmailRequest payload = new EmailRequest
        {
            To = string.Join(";", email.RecipientEmailAddresses),
            Cc = "",
            Subject = Subject,
            Body = await BuildBodyAsync(email, cancellationToken).ConfigureAwait(false),
            IsHtml = true,
        };

        using HttpRequestMessage request = new(HttpMethod.Post, _requestUri);
        request.Headers.Add("x-api-key", _apiKey);
        request.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

        using HttpResponseMessage response = await _httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            string responseBody = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            _logger.LogError(
                "Failed to send legal onboarding email. Status '{StatusCode}'. Body '{Body}'.",
                (int)response.StatusCode,
                responseBody);
            throw new InvalidOperationException(
                $"Graph sendMail failed with status {(int)response.StatusCode}: {responseBody}");
        }

        _logger.LogInformation("Sent legal onboarding email for company '{CompanyName}'.", email.CompanyName);
    }

    private static async Task<string> BuildBodyAsync(
        LegalOnboardingEmailRequest email,
        CancellationToken cancellationToken)
    {
        bool hasNda = !string.IsNullOrWhiteSpace(email.NonDisclosureAgreementNumber);
        string templateName = hasNda
            ? "LegalOnboardingWithNda.html"
            : "LegalOnboardingWithoutNda.html";
        string templatePath = Path.Combine(AppContext.BaseDirectory, "Templates", templateName);
        string body = await File.ReadAllTextAsync(templatePath, cancellationToken).ConfigureAwait(false);

        var replacements = new Dictionary<string, string>
        {
            ["{{CompanyName}}"] = email.CompanyName,
            ["{{CompanyAddress}}"] = email.CompanyAddress,
            ["{{SignatoryName}}"] = email.SignatoryName,
            ["{{SignatoryTitle}}"] = email.SignatoryTitle,
            ["{{SignatoryEmailAddress}}"] = email.SignatoryEmailAddress,
            ["{{NonDisclosureAgreementNumber}}"] = email.NonDisclosureAgreementNumber ?? string.Empty,
        };

        foreach ((string placeholder, string value) in replacements)
        {
            body = body.Replace(placeholder, HtmlEncoder.Default.Encode(value), StringComparison.Ordinal);
        }

        return body;
    }

    private static void Validate(LegalOnboardingEmailRequest email)
    {
        ArgumentNullException.ThrowIfNull(email);
        ArgumentException.ThrowIfNullOrWhiteSpace(email.CompanyName);
        ArgumentException.ThrowIfNullOrWhiteSpace(email.CompanyAddress);
        ArgumentException.ThrowIfNullOrWhiteSpace(email.SignatoryName);
        ArgumentException.ThrowIfNullOrWhiteSpace(email.SignatoryTitle);

        if (email.RecipientEmailAddresses is not { Length: > 0 })
        {
            throw new ArgumentException("At least one recipient email address is required.", nameof(email));
        }

        if (!MailAddress.TryCreate(email.SignatoryEmailAddress, out _)
            || email.RecipientEmailAddresses.Any(address => !MailAddress.TryCreate(address, out _)))
        {
            throw new ArgumentException("All email addresses must be valid.", nameof(email));
        }
    }
}
