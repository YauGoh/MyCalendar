// See https://aka.ms/new-console-template for more information
using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using MyCalendar.Console.Configurations;

var configuration = new ConfigurationBuilder()
     .AddJsonFile($"appsettings.json")
     .Build();

var microsoftGraphConfig = configuration
    .GetRequiredSection("MicrosoftGraph")
    .Get<MicrosoftGraphConfig>();

var scopes = new string[]
{
    $"{microsoftGraphConfig.ApiUrl}.default"
};

var clientSecretCredential = new ClientSecretCredential(
    microsoftGraphConfig.Tenant,
    microsoftGraphConfig.ClientId,
    microsoftGraphConfig.ClientSecret,
    new TokenCredentialOptions
    {
        AuthorityHost = AzureAuthorityHosts.AzurePublicCloud
    });

var graphServiceClient = new GraphServiceClient(clientSecretCredential, scopes);

// Get the user using his Email Address

var users = await graphServiceClient.Users
    .Request()
    .Filter("Mail eq 'yauGohChow@HackathonOct22Definitiv.onmicrosoft.com'")
    .GetAsync();

var user = users.First();

Console.WriteLine($"For the user {user.DisplayName}:");

// Get the Default calendar for the user

var calendars = await graphServiceClient
    .Users[user.Id]
    .Calendars
    .Request()
    .Filter("name eq 'Calendar'")
    .GetAsync();

var calendar = calendars.First();

Console.WriteLine($"    found calendar {calendar.Name} ({calendar.Id})");

// Create the Event

var @event = new Event
{
    Subject = "Annual leave",
    Start = new DateTimeTimeZone
    {
        DateTime = "2022-10-5T00:00:00",
        TimeZone = "W. Australia Standard Time"
    },
    End = new DateTimeTimeZone
    {
        DateTime = "2022-10-6T00:00:00",
        TimeZone = "W. Australia Standard Time"
    },
    IsAllDay = true
};

@event = await graphServiceClient
    .Users[user.Id]
    .Calendars[calendar.Id]
    .Events
    .Request()
    .AddAsync(@event);

Console.WriteLine($"    Event Added {@event.Subject} ({@event.Id})");

Console.ReadLine();