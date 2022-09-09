using DEA;
using WriteLog;
using FolderCleaner;
using Microsoft.Extensions.Configuration;

// TODO 1: Brake the main graph functions into smaller set of chuncks.
// TODO 2: Change the usage of DEA.conf to app.conf (But I don't think it's needed. Bcause I use app.conf to stroe some very important data set.).
// TODO 3: Make the metadata file have the same name as the pdf or attachment file and Remove .pdf extention.
// TODO 4: Make the attachmetn download loop more efficiant. <-- done.
// TODO 5: Check the error folder mover. <-- Working and done.
// TODO 6: Stramline the code.
// TODO 7: Write summeries.
// TODO 8: Create a way to move files to error folder and forward the mail if the file attachments are not accepted.
// TODO 9: Make a internet connection checker.

// Aplication title just for fun.
WriteLogClass.WriteToLog(3, "Starting DEA ....");

// Check for the attachment download folder and the log folder. Then creates the folders if they're missing.
GraphHelper.CheckFolders("none");
WriteLogClass.WriteToLog(3, "Checking main folders ....");

// Clean the main download folder.
FolderCleanerClass.GetFolders(GraphHelper.CheckFolders("Download"));

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
        WriteLogClass.WriteToLog(1, "Set the Graph API permissions. Using dotnet user-secrets set or appsettings.json.... User secrets is not correct.");
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
    WriteLogClass.WriteToLog(1, "Graph client initialization faild  .....");
}
else
{
    WriteLogClass.WriteToLog(3, "Graph client initialization successful ....");
    Thread.Sleep(5000);
    WriteLogClass.WriteToLog(3, "Starting attachment download process ....");
    await GraphHelper.InitializGetAttachment();
}

WriteLogClass.WriteToLog(3, "Email processing ended ...\n");

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