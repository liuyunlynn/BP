namespace MeetingBot;

/// <summary>Information required for the legal onboarding email.</summary>
public sealed class LegalOnboardingEmailRequest
{
    public string CompanyName { get; set; } = string.Empty;

    public string CompanyAddress { get; set; } = string.Empty;

    public string SignatoryName { get; set; } = string.Empty;

    public string SignatoryTitle { get; set; } = string.Empty;

    public string SignatoryEmailAddress { get; set; } = string.Empty;

    public string? NonDisclosureAgreementNumber { get; set; }
}
