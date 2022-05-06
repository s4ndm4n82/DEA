using DEA;
using Microsoft.Extensions.Configuration;

Console.WriteLine("Download Email Attachments (D.E.A)\n");

//check for the attachment download folder and creates the folder if it's missing.
GraphHelper.CheckFolders();

//Getting the Graph and checking the settings for Graph.
var appConfig = LoadAppSettings();

if (appConfig == null)
{
    Console.WriteLine("Set the graph API pemissions. Using dotnet user-secrets set .... They don't exsits in this computer.");
    return;
}

var appId = appConfig["appId"];
var TenantId = appConfig["TenantId"];
var Instance = appConfig["Instance"];
var GraphApiUrl = appConfig["GraphApiUrl"];
//var scopesString = appConfig["scopes"];
//var scopes = scopesString.Split(';');
string[] scopes = new string[] { "https://graph.microsoft.com/.default" };
// Initialize Graph client
/*GraphHelper.Initialize(appId, scopes, (code, cancellation) => {
    Console.WriteLine(code.Message);
    return Task.FromResult(0);
});*/

GraphHelper.InitializeAuto(appId, Instance, TenantId, GraphApiUrl, scopes);

//string? accessToken = GraphHelper.GetAccessTokenAsync(scopes).Result;

int userChoice = -0x1; // Value is -1 in hex.

while (userChoice != 0)
{
    //Console.Clear();
    Console.WriteLine("Please select one of the options from below:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Display Access Token");
    Console.WriteLine("2. Download Attachments");

    try
    {
        userChoice = int.Parse(s: Console.ReadLine());
    }
    catch (System.FormatException)
    {
        // Setting the choice to invalid value.
        userChoice = -0x1; // Value is -1 in hex.
    }

    switch (userChoice)
    {
        case 0:
            Console.Clear();
            Console.WriteLine("\nGood Bye ... !");
            Thread.Sleep(1000);
            Environment.Exit(0);
            break;

        case 1:
            // Display Access Token.
           //Console.WriteLine("Access token: {0}\n", accessToken);
            break;

        case 2:
            // Download the attachments.
            await GraphHelper.GetAttachmentTodayAsync();
            break;

        default:
            Console.WriteLine("Not a valid choice. Try again.");
            break;
    }
}

static IConfigurationRoot? LoadAppSettings()
{
    var appConfig = new ConfigurationBuilder()
        .AddUserSecrets<Program>()
        .Build();

    // Check for required settings
    if (string.IsNullOrEmpty(appConfig["appId"]) ||
        string.IsNullOrEmpty(appConfig["scopes"]))
    {
        return null;
    }

    return appConfig;
}