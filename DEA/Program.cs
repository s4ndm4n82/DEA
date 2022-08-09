﻿using DEA;
using Microsoft.Extensions.Configuration;

// Aplication title just for fun.
Console.WriteLine("Download Email Attachments (D.E.A)\n");

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
        Console.WriteLine("Set the Graph API permissions. Using dotnet user-secrets set .... User secrets is not correct.");
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
    await GraphHelper.InitializGetAttachment();
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