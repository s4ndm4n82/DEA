using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using WriteLog;
using DEA2Levels;
using ReadSettings;
using CreateMetadataFile; // Might need to use this later so leaving it.

namespace DEA
{
    public class GraphHelper
    {
        private static GraphServiceClient? graphClient;        
        private static AuthenticationResult? AuthToken;
        private static IConfidentialClientApplication? Application;

        // Initilize the graph clinet and calls GetAuthTokenWithOutUser() to get the token.
        // If Task<bool> keeps giving an error switch to bool. And change the return Task.FromResult(true) to return true;
        public static Task<bool> InitializeGraphClient(string ClientId, string InstanceId, string TenantId, string GraphUrl, string ClientSecret, string[] scopes)
        {   
            try
            {
                graphClient = new GraphServiceClient(GraphUrl,
                    new DelegateAuthenticationProvider(async (requestMessage) =>
                    {
                        requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", await GetAuthTokenWithOutUser(ClientId, InstanceId, TenantId, ClientSecret, scopes));
                    }
                    ));
                return Task.FromResult(true);
            }
            catch (Exception ex)
            {                
                WriteLogClass.WriteToLog(1, $"Exception at graph client initilizing: {ex.Message}");
                return Task.FromResult(false);
            }
        }

        // Get the token from the Azure according to the default scopes set in the server.
        public static async Task<string> GetAuthTokenWithOutUser(string ClientID, string InstanceID, string TenantID, string ClientSecret, string[] scopes)
        {
            string Authority = string.Concat(InstanceID, TenantID);

            Application = ConfidentialClientApplicationBuilder.Create(ClientID)
                          .WithClientSecret(ClientSecret)
                          .WithAuthority(new Uri(Authority))
                          .Build();
            try
            {
                // Aquirs the token and assigns it AuthToken variable.
                AuthToken = await Application.AcquireTokenForClient(scopes).ExecuteAsync();                
            }
            catch (MsalUiRequiredException ex)
            {
                // The application doesn't have sufficient permissions.
                // - Did you declare enough app permissions during app creation?
                // - Did the tenant admin grant permissions to the application?
                WriteLogClass.WriteToLog(1, $"Exception at token accuire: {ex.Message}");
            }
            catch (MsalServiceException ex) when (ex.Message.Contains("AADSTS70011"))
            {
                // Invalid scope. The scope has to be in the form "https://resourceurl/.default"
                // Mitigation: Change the scope to be as expected.
                WriteLogClass.WriteToLog(1, "Scopes provided are not supported");
            }
            
            return AuthToken!.AccessToken;

        }

        public static async Task InitializGetAttachment()
        {
            // Calls and assigne everything from ReadSettingsClass class.
            var EmailCheckList = new ReadSettingsClass();

            // Loops through all the emails from user accounts property.
            foreach (string Email in EmailCheckList.UserAccounts!)
            {
                // Check emails and match them to execute the correct function.
                try
                {
                    // Regex should match any email address that look like accounting2@efakturamottak.no.
                    Regex EmailRegEx = new Regex(@"^accounting+(?=[0-9]{0,3}@[a-z]+[\.][a-z]{2,3})");
                    if (EmailRegEx.IsMatch(Email))
                    {
                        // Calls the function for reading accounting emails for attachments.
                        await GraphHelper2Levels.GetEmailsAttacmentsAccount(graphClient!, Email);
                    }
                    else
                    {
                        // Calls the function to read ATC emails.
                        await GetAttachmentAtc(Email);
                    }
                }
                catch (Exception ex)
                {
                    WriteLogClass.WriteToLog(1, $"Exception at email loading: {ex.Message}");
                }
            }
        }

        public static async Task GetAttachmentAtc(string _Email) // Looks for 3 levels of folders.
        {
            // Parameters read from the config files.
            var ConfigParam = new ReadSettingsClass();

            bool DateSwitch = ConfigParam.DateFilter;
            int MaxAmountOfEmails = ConfigParam.MaxLoadEmails;
            string ImportFolderPath = Path.Combine(ConfigParam.ImportFolderLetter, ConfigParam.ImportFolderPath);

            var DateToDay = DateTime.Now.ToString("dd.MM.yyyy");

            // TODO: 1. Enable one of the queries before server testing.
            List<QueryOption> SearchOptions;

            if (DateSwitch)
            {
                SearchOptions = new List<QueryOption>
                {
                    new QueryOption("search", $"%22received:{DateToDay}%22")
                    //new QueryOption("search", $"%22hasAttachments:true received:{DateToDay}%22")
                 };
            }
            else
            {
                SearchOptions = new List<QueryOption> { };
            }

            

            try
            {
                /* Can't use .Me with application permissions. It only can be used with delegated permissions.
                 var FirstSubFolderIDs = await graphClient.Me.MailFolders["Inbox"].ChildFolders */

                WriteLogClass.WriteToLog(3, $"Processing Email {_Email}  ....");

                // First level of subfolders under the inbox.
                var FirstSubFolderIDs = await graphClient!.Users[$"{_Email}"].MailFolders["Inbox"].ChildFolders                    
                    .Request()
                    .Select(fid => new
                    {
                        fid.Id,
                        fid.DisplayName
                    })
                    .GetAsync();
                
                foreach(var FirstSubFolderID in FirstSubFolderIDs)
                {
                    if(FirstSubFolderID.Id != null)
                    {
                        // Second level of subfolders under the inbox.
                        var SecondSubFolderIDs = await graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                            .ChildFolders[$"{FirstSubFolderID.Id}"]
                            .ChildFolders
                            .Request()
                            .Select(sid => new
                            {
                                sid.Id,
                                sid.DisplayName
                            })
                            .GetAsync();

                        //foreach (var SecondSubFolderID in SecondSubFolderIDs.Where(x => string.IsNullOrWhiteSpace(x.Id))) <-- use this to cleaup the fif switches.
                        foreach (var SecondSubFolderID in SecondSubFolderIDs)
                        {
                            if(SecondSubFolderID.Id != null)
                            {
                                // Third level of subfolders under the inbox.
                                var ThirdSubFolderIDs = await graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                                    .ChildFolders[$"{FirstSubFolderID.Id}"]
                                    .ChildFolders[$"{SecondSubFolderID.Id}"]
                                    .ChildFolders
                                    .Request()
                                    .Select(tid => new
                                    {
                                        tid.Id,
                                        tid.DisplayName,
                                    })
                                    .GetAsync();               

                                foreach (var ThirdSubFolderID in ThirdSubFolderIDs)
                                {
                                    // Getting emails from the last child folder.
                                    if (ThirdSubFolderID.DisplayName == "Processing")
                                    {
                                        WriteLogClass.WriteToLog(3, $"Processing folder path {FirstSubFolderID.DisplayName} -> {SecondSubFolderID.DisplayName} -> {ThirdSubFolderID.DisplayName}");
                                        // Looping through the emails in the subfolder "Processing".
                                        var GetMessageAttachments = await graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                                            .ChildFolders[$"{FirstSubFolderID.Id}"]
                                            .ChildFolders[$"{SecondSubFolderID.Id}"]
                                            .ChildFolders[$"{ThirdSubFolderID.Id}"]
                                            .Messages
                                            .Request(SearchOptions)
                                            .Expand("attachments")
                                            .Select(gma => new
                                            {   
                                                gma.Id,
                                                gma.Subject,
                                                gma.HasAttachments,                                               
                                                gma.Attachments
                                            })
                                            .Top(MaxAmountOfEmails)
                                            .GetAsync();

                                        // Gets the message count.
                                        var MessageCount = GetMessageAttachments.Count;

                                        if (MessageCount != 0)
                                        {
                                            // Loops through the emails.
                                            foreach (var Message in GetMessageAttachments)
                                            {
                                                // Check if the mail has an attachment or not.
                                                var HasFileAttached = Message.HasAttachments;
                                                
                                                // If varible is true the attachment will be downloaded.
                                                if (HasFileAttached == true)
                                                {
                                                    // Counting the aount of messages with attachments. To loop through below.
                                                    var AttachmentCount = Message.Attachments.Count;

                                                    // Assigning display names.                                            
                                                    var FirstFolderName = FirstSubFolderID.DisplayName;
                                                    var SecondFolderName = SecondSubFolderID.DisplayName;

                                                    // Creating the destnation folders.
                                                    string[] MakeDestinationFolderPath = { ImportFolderPath, _Email, FirstFolderName, SecondFolderName };
                                                    var DestinationFolderPath = Path.Combine(MakeDestinationFolderPath);

                                                    // Variable used to store all the accepted extentions.
                                                    // string[] AcceptedExtention = { ".pdf", ".jpg" }; // Make this read from text file.
                                                    string[] AcceptedExtention = ConfigParam.AllowedExtentions;

                                                    // Initilizing the download folder path variable.
                                                    string PathFullDownloadFolder = string.Empty;

                                                    // Used to set the status of if the email is moved or not.
                                                    // If moved this will be true. If not the code for forwarding will be executed.
                                                    var EmailMoveStatus = true;

                                                    // For loop to go through all the extentions from extentions variable.
                                                    for (int i = 0; i < AcceptedExtention.Length; ++i)
                                                    {
                                                        // Select all the items from attachments variable that contains matching extentions.
                                                        var AcceptedExtensionCollection = Message.Attachments.Where(x => x.Name.ToLower().EndsWith(AcceptedExtention[i]));

                                                        // Checks the collection empty or not.
                                                        if (AcceptedExtensionCollection.Any(y => y.Name.ToLower().Contains(AcceptedExtention[i])))
                                                        {
                                                            WriteLogClass.WriteToLog(3, "Collection check succeeded ...");

                                                            // FolderNameRnd creates a 10 digit folder name. CheckFolder returns the download path.
                                                            // This has to be called here. Don't put it within the for loop or it will start calling this
                                                            // each time and make folder for each file. Also calling this out side of the extentions FOR loop.
                                                            // causes an exception error at the "DownloadFileExistTest" test due file not been available.
                                                            PathFullDownloadFolder = Path.Combine(CheckFolders("Download"), FolderNameRnd(10));


                                                            foreach (var Attachment in AcceptedExtensionCollection)
                                                            {
                                                                // Switch to execute the atachment download.
                                                                // If this is false that means the message has been moved.
                                                                bool MsgSwitch = true;

                                                                try
                                                                {
                                                                    // Check if the selected email message exsits or not.
                                                                    // If not exception will be thrown which will be captured in the catch area and it sets the MsgSwitch to false.
                                                                    // Which will make the next if skip the attachment download. This is done to avoide the error that occurs from
                                                                    // having signaturs as attachments.
                                                                    var CheckMsgId = await graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                                                                                .ChildFolders[$"{FirstSubFolderID.Id}"]
                                                                                .ChildFolders[$"{SecondSubFolderID.Id}"]
                                                                                .ChildFolders[$"{ThirdSubFolderID.Id}"]
                                                                                .Messages[$"{Message.Id}"]
                                                                                .Request()
                                                                                .GetAsync();
                                                                }
                                                                catch (ServiceException ex)
                                                                {
                                                                    if (ex.Error.Code == "ErrorItemNotFound")
                                                                    {
                                                                        MsgSwitch = false;
                                                                    }
                                                                }

                                                                if (MsgSwitch)
                                                                {
                                                                    WriteLogClass.WriteToLog(3, $"Starting attachment download from {Message.Subject} ....");

                                                                    // Should mark and download the itemattacment which is the correct attachment.
                                                                    var TrueAttachment = await graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                                                                                    .ChildFolders[$"{FirstSubFolderID.Id}"]
                                                                                    .ChildFolders[$"{SecondSubFolderID.Id}"]
                                                                                    .ChildFolders[$"{ThirdSubFolderID.Id}"]
                                                                                    .Messages[$"{Message.Id}"]
                                                                                    .Attachments[$"{Attachment.Id}"]
                                                                                    .Request()
                                                                                    .Expand("microsoft.graph.itemattachment/item")
                                                                                    .GetAsync();

                                                                    // Details of the attachment.
                                                                    var TrueAttachmentProps = (FileAttachment)TrueAttachment;
                                                                    string TrueAttachmentName = TrueAttachmentProps.Name;
                                                                    byte[] TruAttachmentBytes = TrueAttachmentProps.ContentBytes;
                                                                    
                                                                    // Saves the file to the local hard disk.
                                                                    DownloadAttachedFiles(PathFullDownloadFolder, TrueAttachmentName, TruAttachmentBytes);

                                                                    WriteLogClass.WriteToLog(3, $"Downloaded attachments from {Message.Subject}   ....");

                                                                    // Creating the metdata file.
                                                                    //var FileFlag = CreateMetaDataXml.GetToEmail4Xml(graphClient, FirstSubFolderID.Id, SecondSubFolderID.Id, ThirdSubFolderID.Id, Message.Id, _Email, PathFullDownloadFolder, TrueAttachmentName);
                                                                    var FileFlag = true;

                                                                    // Directory and file existence check. If not exists it will not return anything.
                                                                    string[] DownloadFolderExistTest = System.IO.Directory.GetDirectories(GraphHelper.CheckFolders("Download")); // Use the main path not the entire download path
                                                                    string[] DownloadFileExistTest = System.IO.Directory.GetFiles(PathFullDownloadFolder); // This causs an erro when the file is not there.

                                                                    if (DownloadFolderExistTest.Length != 0 && DownloadFileExistTest.Length != 0 && FileFlag)
                                                                    {
                                                                        WriteLogClass.WriteToLog(3, "Moving downloaded files to local folder ....");
                                                                        // Moves the downloaded files to destination folder. This would create the folder path if it's missing.
                                                                        if (MoveFolder(PathFullDownloadFolder, DestinationFolderPath))
                                                                        {
                                                                            WriteLogClass.WriteToLog(3, "File moved successfully ....");
                                                                            // Search option sets the $filter query to only get the folder named downloaded.
                                                                            var FolderSearchOption = new List<QueryOption>
                                                                            {
                                                                                new QueryOption ("filter", $"displayName eq %27Exported%27")
                                                                            };

                                                                            try
                                                                            {
                                                                                // Loop through and selects only the Exported folder.
                                                                                var DestinationDetails = await graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                                                                                    .ChildFolders[$"{FirstSubFolderID.Id}"]
                                                                                    .ChildFolders[$"{SecondSubFolderID.Id}"]
                                                                                    .ChildFolders
                                                                                    .Request(FolderSearchOption)
                                                                                    .GetAsync();

                                                                                foreach (var Destination in DestinationDetails)
                                                                                {
                                                                                    if (Destination.DisplayName == "Exported") // Just a backup check of the folder name.
                                                                                    {
                                                                                        var MessageID = Message.Id;
                                                                                        var MoveDestinationID = Destination.Id;

                                                                                        // Moves the mail to downloaded folder.
                                                                                        if (await GraphHelper.MoveEmails(FirstSubFolderID.Id, SecondSubFolderID.Id, ThirdSubFolderID.Id, MessageID, MoveDestinationID, _Email))
                                                                                        {
                                                                                            WriteLogClass.WriteToLog(3, $"Email {Message.Subject} moved to export folder ...");
                                                                                            EmailMoveStatus = false;
                                                                                        }
                                                                                        else
                                                                                        {
                                                                                            WriteLogClass.WriteToLog(3, $"Email {Message.Subject} is not moved to export folder ...");
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }
                                                                            catch (Exception ex)
                                                                            {
                                                                                WriteLogClass.WriteToLog(1, $"Exception at attachment download area: {ex.Message}");
                                                                            }

                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if (EmailMoveStatus) // Executes if variable is false.
                                                            {
                                                                // Search for the subfolder named error.
                                                                var FolderSearchOption2 = new List<QueryOption>
                                                                {
                                                                    new QueryOption ("filer", $"displayName eq %27Error%27")
                                                                };

                                                                //Loop thorugh to Select only error folder from the subfolders.
                                                                var ErroFolderDetails = await graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                                                                    .ChildFolders[$"{FirstSubFolderID.Id}"]
                                                                    .ChildFolders[$"{SecondSubFolderID.Id}"]
                                                                    .ChildFolders
                                                                    .Request(FolderSearchOption2)
                                                                    .GetAsync();

                                                                foreach (var ErrorFolder in ErroFolderDetails)
                                                                {
                                                                    if (ErrorFolder.DisplayName == "Error") // Just a backup check of the folder name.
                                                                    {
                                                                        // Folder ID and the message ID that need to be forwarded to the client.
                                                                        var MessageID2 = Message.Id;
                                                                        var ErrorFolderId = ErrorFolder.Id;
                                                                        var AttachmentStatus = 1; // Extension not accepted.

                                                                        // Email is beeing forwarded.
                                                                        var ForwardDone = await GraphHelper.ForwardEmtpyEmail(FirstSubFolderID.Id, SecondSubFolderID.Id, ErrorFolderId, MessageID2, _Email, AttachmentStatus);

                                                                        // After forwarding checks if the action returned true.
                                                                        // Item2 is the bool value returned.
                                                                        // Item1 is the maile address.
                                                                        if (ForwardDone.Item2)
                                                                        {
                                                                            WriteLogClass.WriteToLog(3, $"Email forwarded to {ForwardDone.Item1} ....");
                                                                        }
                                                                        else
                                                                        {
                                                                            WriteLogClass.WriteToLog(3, $"Email not forwarded to {ForwardDone.Item1} ...");
                                                                        }

                                                                        // Moves the empty emails to error folder once forwarding is done.
                                                                        if (await GraphHelper.MoveEmails(FirstSubFolderID.Id, SecondSubFolderID.Id, ThirdSubFolderID.Id, MessageID2, ErrorFolderId, _Email))
                                                                        {
                                                                            WriteLogClass.WriteToLog(3, $"Mail Moved to {ErrorFolder.DisplayName} Folder ....");
                                                                        }
                                                                        else
                                                                        {
                                                                            WriteLogClass.WriteToLog(3, $"Mail was Not Moved to {ErrorFolder.DisplayName} Folder ....");
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (HasFileAttached == false)
                                                    {
                                                        // Search for the subfolder named error.
                                                        var FolderSearchOption2 = new List<QueryOption>
                                                        {
                                                            new QueryOption ("filer", $"displayName eq %27Error%27")
                                                        };

                                                        //Loop thorugh to Select only error folder from the subfolders.
                                                        try
                                                        {
                                                            var ErroFolderDetails = await graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                                                            .ChildFolders[$"{FirstSubFolderID.Id}"]
                                                            .ChildFolders[$"{SecondSubFolderID.Id}"]
                                                            .ChildFolders
                                                            .Request(FolderSearchOption2)
                                                            .GetAsync();

                                                            foreach (var ErrorFolder in ErroFolderDetails)
                                                            {
                                                                if (ErrorFolder.DisplayName == "Error") // Just a backup check of the folder name.
                                                                {
                                                                    // Folder ID and the message ID that need to be forwarded to the client.
                                                                    var MessageID2 = Message.Id;
                                                                    var ErrorFolderId = ErrorFolder.Id;
                                                                    var AttachmentStatus = 0;

                                                                    // Email is beeing forwarded.
                                                                    var ForwardDone = await ForwardEmtpyEmail(FirstSubFolderID.Id, SecondSubFolderID.Id, ErrorFolderId, MessageID2, _Email, AttachmentStatus);

                                                                    // After forwarding checks if the action returned true.
                                                                    // Item2 is the bool value returned.
                                                                    // Item1 is the maile address.
                                                                    if (ForwardDone.Item2)
                                                                    {
                                                                        WriteLogClass.WriteToLog(3, $"Email forwarded to {ForwardDone.Item1} ....");
                                                                    }
                                                                    else
                                                                    {
                                                                        WriteLogClass.WriteToLog(3, $"Email not forwarded to {ForwardDone.Item1} ....");
                                                                    }

                                                                    // Moves the empty emails to error folder once forwarding is done.
                                                                    if (await GraphHelper.MoveEmails(FirstSubFolderID.Id, SecondSubFolderID.Id, ThirdSubFolderID.Id, MessageID2, ErrorFolderId, _Email))
                                                                    {
                                                                        WriteLogClass.WriteToLog(3, $"Mail Moved to {ErrorFolder.DisplayName} Folder ....");
                                                                    }
                                                                    else
                                                                    {
                                                                        WriteLogClass.WriteToLog(3, $"Mail was Not Moved to {ErrorFolder.DisplayName} Folder ....");
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        catch (Exception ex)
                                                        {
                                                            WriteLogClass.WriteToLog(1, $"Exception at error folder mover: {ex.Message}");
                                                        }                                                        
                                                    }
                                                }
                                            }
                                        }                                        
                                    }                             
                                }
                            }                            
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLogClass.WriteToLog(1, $"Exception at end of main foreach: {ex.Message}");
            }
        }

        // Generate the random 10 digit number as the folder name.
        public static string FolderNameRnd(int RndLength)
        {
            Random RndNumber = new();
            string NumString = string.Empty;
            for(int i = 0; i < RndLength; i++)
            {
                NumString = String.Concat(NumString, RndNumber.Next(10).ToString());
            }
            return NumString;
        }

        // Check the exsistance of the download folders.
        public static string CheckFolders(string FolderSwitch)
        {
            // Get current execution path.
            string FolderPath = string.Empty;
            string? PathRootFolder = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string DownloadFolderName = "Download";
            string LogFolderName = "Logs";
            string PathDownloadFolder = Path.Combine(PathRootFolder!, DownloadFolderName);
            string PathLogFolder = Path.Combine(PathRootFolder!, LogFolderName);

            // Check if download folder exists. If not creates the fodler.
            if (!System.IO.Directory.Exists(PathDownloadFolder))
            {
                try
                {
                    System.IO.Directory.CreateDirectory(PathDownloadFolder);                    
                }
                catch (Exception ex)
                {
                    WriteLogClass.WriteToLog(1, $"Exception at download folder creation: {ex.Message}");
                }
            }

            if (!System.IO.Directory.Exists(PathLogFolder))
            {
                try
                {
                    System.IO.Directory.CreateDirectory(PathLogFolder);                    
                }
                catch (Exception ex)
                {
                    WriteLogClass.WriteToLog(1, $"Exception at download folder creation: {ex.Message}");
                }
            }

            if (FolderSwitch == "Download")
            {
                FolderPath = PathDownloadFolder;
            }
            else if (FolderSwitch == "Log")
            {
                FolderPath = PathLogFolder;
            }
            else
            {
                FolderPath = string.Empty;
            }

            return FolderPath;
        }

        // Downnloads the attachments to local harddrive.
        public static bool DownloadAttachedFiles(string DownloadFolderPath, string DownloadFileName, byte[] DownloadFileData)
        {
            if (!System.IO.Directory.Exists(DownloadFolderPath))
            {
                try
                {
                    System.IO.Directory.CreateDirectory(DownloadFolderPath);
                }
                catch (Exception ex)
                {
                    WriteLogClass.WriteToLog(1, $"Exception at download folder creation: {ex.Message}");
                }
            }

            try
            {
                // Full path for the attachment to be downloaded with the attachment name
                var PathFullDownloadFile = Path.Combine(DownloadFolderPath, DownloadFileName);
                System.IO.File.WriteAllBytes(PathFullDownloadFile, DownloadFileData);
                return true;
            }
            catch (Exception ex)
            {
                WriteLogClass.WriteToLog(1, $"Exception at download path: {ex.Message}");
                return false;
            }
        }

        // Move the folder to main import folder on the local machine.
        public static bool MoveFolder(string SourceFolderPath, string DestiFolderPath)
        {
            try
            {  
                var SourceLastFolder = SourceFolderPath.Split(Path.DirectorySeparatorChar).Last();
                var FullDestinationPath = Path.Combine(DestiFolderPath, SourceLastFolder);
                Microsoft.VisualBasic.FileIO.FileSystem.MoveDirectory(SourceFolderPath, FullDestinationPath);
                return true;
            }
            catch (Exception ex)
            {
                WriteLogClass.WriteToLog(1, $"Error getting event: {ex.Message}");
                return false;
            }
        }

        // Forwards emails with out any attachment to the sender.
        public static async Task<(string?,bool)> ForwardEmtpyEmail(string FolderId1, string FolderId2, string ErrFolderId, string MsgId2, string _Email, int AttnStatus)
        {
            try
            {
                if (!string.IsNullOrEmpty(FolderId2))
                {
                    // Get ths the emails details like subject and from email.
                    var MsgDetails = await graphClient!.Users[$"{_Email}"].MailFolders["Inbox"]
                            .ChildFolders[$"{FolderId1}"]
                            .ChildFolders[$"{FolderId2}"]
                            .ChildFolders[$"{ErrFolderId}"]
                            .Messages[$"{MsgId2}"]
                            .Request()
                            .Select(em => new
                            {
                                em.Subject,
                                em.From
                            })
                            .GetAsync();

                    // Variables to be used with graph forward.
                    var FromName    = MsgDetails.From.EmailAddress.Name;
                    var FromEmail   = MsgDetails.From.EmailAddress.Address;
                    var MailComment = string.Empty;

                    if (AttnStatus != 1)
                    {
                        MailComment = "Hi,<br />" +
                            "<b>Please don't respond to this email. We're testing a new system.</b><br />" +
                            "You might see a few emails like this, just ignore them. Sorry for the inconvenience.";
                        //"Hi,<br /> Below email doesn't contain any attachment."; // Can be change with html.
                    }
                    else
                    {
                        MailComment = "Hi,<br />" +
                            "<b>Please don't respond to this email. We're testing a new system.</b><br />" +
                            "You might see a few emails like this, just ignore them. Sorry for the inconvenience.";
                        //"Below email's attachment file type is not accepted. Please send attachments as .pdf or .jpg.";
                    }

                    // Recipient setup for the mail header.
                    var toRecipients = new List<Recipient>()
                    {
                        new Recipient
                        {
                            EmailAddress = new EmailAddress
                            {
                                Name = FromName,
                                Address = FromEmail
                            }
                        }
                    };

                    // Forwarding the non attachment email using .forward().
                    await graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                        .ChildFolders[$"{FolderId1}"]
                        .ChildFolders[$"{FolderId2}"]
                        .ChildFolders[$"{ErrFolderId}"]
                        .Messages[$"{MsgId2}"]
                        .Forward(toRecipients, null, MailComment)
                        .Request()
                        .PostAsync();

                    return (FromEmail, true);
                }
                else
                {
                    // Get ths the emails details like subject and from email.
                    var MsgDetails = await graphClient!.Users[$"{_Email}"].MailFolders["Inbox"]
                            .ChildFolders[$"{FolderId1}"]
                            .ChildFolders[$"{ErrFolderId}"]
                            .Messages[$"{MsgId2}"]
                            .Request()
                            .Select(em => new
                            {
                                em.Subject,
                                em.From
                            })
                            .GetAsync();

                    // Variables to be used with graph forward.
                    var FromName = MsgDetails.From.EmailAddress.Name;
                    var FromEmail = MsgDetails.From.EmailAddress.Address;
                    var MailComment = string.Empty;

                    if (AttnStatus != 1)
                    {
                        MailComment = "Hi,<br />" +
                            "<b>Please don't respond to this email. We're testing a new system.</b><br />" +
                            "You might see a few emails like this, just ignore them. Sorry for the inconvenience.";
                        //"Hi,<br /> Below email doesn't contain any attachment."; // Can be change with html.
                    }
                    else
                    {
                        MailComment = "Hi,<br />" +
                            "<b>Please don't respond to this email. We're testing a new system.</b><br />" +
                            "You might see a few emails like this, just ignore them. Sorry for the inconvenience.";
                        //"Below email's attachment file type is not accepted. Please send attachments as .pdf or .jpg.";
                    }

                    // Recipient setup for the mail header.
                    var toRecipients = new List<Recipient>()
                {
                    new Recipient
                    {
                        EmailAddress = new EmailAddress
                        {
                            Name = FromName,
                            Address = FromEmail
                        }
                    }
                };

                    // Forwarding the non attachment email using .forward().
                    await graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                        .ChildFolders[$"{FolderId1}"]
                        .ChildFolders[$"{ErrFolderId}"]
                        .Messages[$"{MsgId2}"]
                        .Forward(toRecipients, null, MailComment)
                        .Request()
                        .PostAsync();

                    return (FromEmail, true);
                }
                
            }
            catch (Exception ex)
            {
                return (ex.Message, false);
            }
        }

        //Moves the email to Downloded folder.
        public static async Task<bool> MoveEmails(string FirstFolderId, string SecondFolderId, string ThirdFolderId, string MsgId, string DestiId, string _Email)
        {
            try
            {
                if (!string.IsNullOrEmpty(ThirdFolderId))
                {
                    //Graph api call to move the email message.
                    await graphClient!.Users[$"{_Email}"].MailFolders["Inbox"]
                        .ChildFolders[$"{FirstFolderId}"]
                        .ChildFolders[$"{SecondFolderId}"]
                        .ChildFolders[$"{ThirdFolderId}"]
                        .Messages[$"{MsgId}"]
                        .Move(DestiId)
                        .Request()
                        .PostAsync();
                }
                else
                {
                    //Graph api call to move the email message.
                    await graphClient!.Users[$"{_Email}"].MailFolders["Inbox"]
                        .ChildFolders[$"{FirstFolderId}"]
                        .ChildFolders[$"{SecondFolderId}"]
                        .Messages[$"{MsgId}"]
                        .Move(DestiId)
                        .Request()
                        .PostAsync();
                }
                return true;
            }
            catch (Exception ex)
            {
                WriteLogClass.WriteToLog(1, $"Exception at moving emails to folders: {ex.Message}");
                return false;
            }
        }
    }
}