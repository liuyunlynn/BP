namespace MeetingBot;

/// <summary>
/// Strongly-typed configuration for legal onboarding email sender identity.
/// Bind from the "LegalOnboardingEmail" section of appsettings.json / user-secrets / environment variables.
/// </summary>
public sealed class LegalOnboardingEmailOptions
{
    /// <summary>API key for the Power Automate email gateway. Supports appsettings.json and App Service environment variables.</summary>
    public string ApiKey { get; set; } = string.Empty;

    /// <summary>Power Automate workflow invocation URL for the legal onboarding email flow.</summary>
    public string RequestUri { get; set; } = string.Empty;
}
