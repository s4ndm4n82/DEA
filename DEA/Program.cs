using DEA;
using Microsoft.Extensions.Configuration;

Console.WriteLine("Download Email Attachments (D.E.A)\n");

var appConfig = LoadAppSettings();

if (appConfig == null)
{
    Console.WriteLine("Missing or invalid appsettings.json...exiting");
    return;
}

var appId = appConfig["appId"];
var scopesString = appConfig["scopes"];
var scopes = scopesString.Split(';');

// Initialize Graph client
GraphHelper.Initialize(appId, scopes, (code, cancellation) => {
    Console.WriteLine(code.Message);
    return Task.FromResult(0);
});

string? accessToken = GraphHelper.GetAccessTokenAsync(scopes).Result;

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
            Console.WriteLine("Access token: {0}\n", accessToken);
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

/*static async void ListAttachments()
{
    var MailMessages = GraphHelper.GetAttachmentToday().Result;

    Console.WriteLine("Attacments:");
    Console.WriteLine($"{MailMessages[0]}");
    Console.WriteLine("\n***************************\n");
    foreach (var Message in MailMessages)
    {
        Console.WriteLine("ID : {0}", Message.Id);
        Console.WriteLine("Display Name : {0}", Message.DisplayName);     

        foreach (var Attachment in Message.Attachments)
        {
            var Item = (FileAttachment)Attachment;
            var Folder = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            var FilePath = Path.Combine(Folder, Item.Name);
            System.IO.File.WriteAllBytes(FilePath, Item.ContentBytes);
        } 

    }

    await GraphHelper.GetAttachmentTodayAsync();
}*/