using DEA;
using Microsoft.Extensions.Configuration;
using ReadSettings;

// Aplication title just for fun.
Console.WriteLine("Download Email Attachments (D.E.A)\n");

/*var Test = new ReadSettingsClass();
Console.WriteLine("Date Filter Switch: {0}", Test.DateFilter);
Console.WriteLine("Max emails to Load: {0}", Test.MaxLoadEmails);

foreach (var Email in Test.UserAccounts)
{
    Console.WriteLine("Email List: {0}", Email);
}

Console.WriteLine("Import Folder Letter: {0}", Test.ImportFolderLetter);
Console.WriteLine("Import Folder Path: {0}", Test.ImportFolderPath);

foreach (var ext in Test.AllowedExtentions)
{
    Console.WriteLine("Allowed Extentions: {0}", ext);
}*/

// Check for the attachment download folder and creates the folder if it's missing.
GraphHelper.CheckFolders();

// Getting the Graph and checking the settings for Graph.
var appConfig = LoadAppSettings();

// Declaring variable to be used with in the if below.
var ClientId = string.Empty;
var TenantId = string.Empty;
var Instance = string.Empty;
var GraphApiUrl = string.Empty;
var ClientSecret = string.Empty;
string[] Scopes = new string[] { };

// If appConfig is equal to null look for settings with in the appsettings.json file.
if (appConfig == null)
{
    // Read the appsettings json file and loads the text in to AppCofigJson variable.
    // File should be with in the main working directory.
    var AppConfigJson = new ConfigurationBuilder()
        .SetBasePath(Directory.GetCurrentDirectory())
        .AddJsonFile("appsettings.json", optional: true, reloadOnChange: true)
        .Build();

    // Initilize the variables with values.
    ClientId = AppConfigJson.GetSection("GraphConfig").GetSection("ClientId").Value;
    TenantId = AppConfigJson.GetSection("GraphConfig").GetSection("TenantId").Value;
    Instance = AppConfigJson.GetSection("GraphConfig").GetSection("Instance").Value;
    GraphApiUrl = AppConfigJson.GetSection("GraphConfig").GetSection("GraphApiUrl").Value;
    ClientSecret = AppConfigJson.GetSection("GraphConfig").GetSection("ClientSecret").Value;
    Scopes = new string[] { $"{AppConfigJson.GetSection("GraphConfig").GetSection("Scopes").Value}" };

    // If Json file is also returns empty then below error would be shown.
    if (string.IsNullOrEmpty(ClientId) ||
        string.IsNullOrEmpty(TenantId) ||
        string.IsNullOrEmpty(Instance) ||
        string.IsNullOrEmpty(GraphApiUrl) ||
        string.IsNullOrEmpty(ClientSecret))
    {
        Console.WriteLine("Set the Graph API permissions. Using dotnet user-secrets set or appsettings.json.... User secrets is not correct.");
    }
}
else
{
    // If appConfig is not equal to null then assings all the setting to variables from UserSecrets.
    ClientId = appConfig["ClientId"];
    TenantId = appConfig["TenantId"];
    Instance = appConfig["Instance"];
    GraphApiUrl = appConfig["GraphApiUrl"];
    ClientSecret = appConfig["ClientSecret"];
    Scopes = new string[] { $"{appConfig["Scopes"]}" };// Gets the application permissions which are set from the Azure AD.
}

// Calls InitializeGraphClient to get the token and connect to the graph API.
if (!await GraphHelper.InitializeGraphClient(ClientId, Instance, TenantId, GraphApiUrl, ClientSecret, Scopes))
{
    Console.WriteLine("Graph client initialization faild.");
}
else
{
    Console.WriteLine("Graph client initialization successful.");
    Console.WriteLine(Environment.NewLine);
    Console.WriteLine("Starting attachment download process.");
    Console.WriteLine(Environment.NewLine);
    await GraphHelper.InitializGetAttachment();
}

// Loads the settings from user sectrets file.
static IConfigurationRoot? LoadAppSettings()
{
    var appConfigUs = new ConfigurationBuilder()
         .AddUserSecrets<Program>()
         .Build();

     // Check for required settings in app secrets.
     if (string.IsNullOrEmpty(appConfigUs["ClientId"]) ||
         string.IsNullOrEmpty(appConfigUs["TenantId"]) ||
         string.IsNullOrEmpty(appConfigUs["Instance"]) ||
         string.IsNullOrEmpty(appConfigUs["GraphApiUrl"]) ||
         string.IsNullOrEmpty(appConfigUs["ClientSecret"]))
     {
         return null;
     }
     else
     {
         return appConfigUs;
     }
}