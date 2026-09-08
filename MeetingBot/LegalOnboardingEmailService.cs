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
    private const string AyanaEmail = "amcginnis@microsoft.com";
    private const string PartnerProgramEmail = "TeamsCategoryPartner@microsoft.com";
    private static readonly string[] GraphScope = ["https://graph.microsoft.com/.default"];

    private readonly BotOptions _options;
    private readonly HttpClient _httpClient;
    private readonly ClientSecretCredential _credential;
    private readonly ILogger<LegalOnboardingEmailService> _logger;

    public LegalOnboardingEmailService(
        BotOptions options,
        IHttpClientFactory httpClientFactory,
        ILogger<LegalOnboardingEmailService> logger)
    {
        _options = options;
        _httpClient = httpClientFactory.CreateClient(nameof(LegalOnboardingEmailService));
        _credential = new ClientSecretCredential(options.TenantId, options.AppId, options.AppSecret);
        _logger = logger;
    }

    public async Task SendAsync(LegalOnboardingEmailRequest email, CancellationToken cancellationToken)
    {
        Validate(email);

        AccessToken token = await _credential
            .GetTokenAsync(new TokenRequestContext(GraphScope), cancellationToken)
            .ConfigureAwait(false);

        string senderUserId = string.IsNullOrWhiteSpace(_options.EmailSenderUserId)
            ? _options.OrganizerUserId
            : _options.EmailSenderUserId;
        if (string.IsNullOrWhiteSpace(senderUserId))
        {
            throw new InvalidOperationException("Bot:EmailSenderUserId or Bot:OrganizerUserId must be configured.");
        }

        object payload = new
        {
            message = new
            {
                subject = Subject,
                body = new
                {
                    contentType = "HTML",
                    content = await BuildBodyAsync(email, cancellationToken).ConfigureAwait(false),
                },
                toRecipients = new[]
                {
                    Recipient(email.SignatoryEmailAddress),
                },
                ccRecipients = new[]
                {
                    Recipient(AyanaEmail),
                    Recipient(PartnerProgramEmail),
                },
            },
            saveToSentItems = true,
        };

        string requestUri = $"https://graph.microsoft.com/v1.0/users/{Uri.EscapeDataString(senderUserId)}/sendMail";
        using HttpRequestMessage request = new(HttpMethod.Post, requestUri);
        request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", token.Token);
        request.Content = new StringContent(JsonSerializer.Serialize(payload), Encoding.UTF8, "application/json");

        using HttpResponseMessage response = await _httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false);
        if (!response.IsSuccessStatusCode)
        {
            string responseBody = await response.Content.ReadAsStringAsync(cancellationToken).ConfigureAwait(false);
            _logger.LogError(
                "Failed to send legal onboarding email. Status '{StatusCode}'. Body '{Body}'.",
                (int)response.StatusCode,
                responseBody);
            throw new InvalidOperationException($"Graph sendMail failed with status {(int)response.StatusCode}.");
        }

        _logger.LogInformation("Sent legal onboarding email for company '{CompanyName}'.", email.CompanyName);
    }

    private static object Recipient(string address) => new
    {
        emailAddress = new
        {
            address,
        },
    };

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

        if (!MailAddress.TryCreate(email.SignatoryEmailAddress, out _))
        {
            throw new ArgumentException("A valid signatory email address is required.", nameof(email));
        }
    }
}
