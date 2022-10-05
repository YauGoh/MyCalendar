namespace MyCalendar.Console.Configurations;

public class MicrosoftGraphConfig
{
    public string Instance { get; init; } = "https://login.microsoftonline.com/{0}";

    public string ApiUrl { get; init; } = "https://graph.microsoft.com/";

    public string Tenant { get; init; } = "consumers";

    public string ClientId { get; init; } = string.Empty;

    public string ClientSecret { get; init; } = string.Empty;

    public string Authority => string.Format(Instance, Tenant);
}
