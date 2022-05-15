using DEA;
using Microsoft.Extensions.Configuration;

// Aplication title just for fun.
Console.WriteLine("Download Email Attachments (D.E.A)\n");

// Check for the attachment download folder and creates the folder if it's missing.
GraphHelper.CheckFolders();

// Getting the Graph and checking the settings for Graph.
var appConfig = LoadAppSettings();

// If appConfig returns null message is shown.
if (appConfig == null)
{
    Console.WriteLine("Set the Graph API permissions. Using dotnet user-secrets set .... User secrets is not correct.");
    return;
}

// Assings all the setting to variables.
var ClientId = appConfig["ClientId"];
var TenantId = appConfig["TenantId"];
var Instance = appConfig["Instance"];
var GraphApiUrl = appConfig["GraphApiUrl"];
var ClientSecret = appConfig["ClientSecret"];
string[] scopes = new string[] { "https://graph.microsoft.com/.default" };// Gets the application permissions which are set from the Azure AD.

// Calls InitializeGraphClient to get the token and connect to the graph API.
if (!await GraphHelper.InitializeGraphClient(ClientId, Instance, TenantId, GraphApiUrl, ClientSecret, scopes))
{
    Console.WriteLine("Graph client initialization faild.");
}
else
{
    Console.WriteLine("Graph client initialization successful.");
    Console.WriteLine(Environment.NewLine);
    await GraphHelper.GetAttachmentTodayAsync();
}

/*int userChoice = -0x1; // Value is -1 in hex.

while (userChoice != 0)
{
    //Console.Clear();
    Console.WriteLine("Please select one of the options from below:");
    Console.WriteLine("0. Exit");
    Console.WriteLine("1. Download Attachments");

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
            // Download the attachments.
            await GraphHelper.GetAttachmentTodayAsync();
            break;

        default:
            Console.WriteLine("Not a valid choice. Try again.");
            break;
    }
}*/

// Loads the settings from user sectrets file.
static IConfigurationRoot? LoadAppSettings()
{
    var appConfig = new ConfigurationBuilder()
        .AddUserSecrets<Program>()
        .Build();

    // Check for required settings in app secrets.
    if (string.IsNullOrEmpty(appConfig["ClientId"])||
        string.IsNullOrEmpty(appConfig["TenantId"])||
        string.IsNullOrEmpty(appConfig["Instance"])||
        string.IsNullOrEmpty(appConfig["GraphApiUrl"])||
        string.IsNullOrEmpty(appConfig["TenantId"])||
        string.IsNullOrEmpty(appConfig["ClientSecret"]))
    {
        return null;
    }

    return appConfig;
}