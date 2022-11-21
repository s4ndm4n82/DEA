using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using WriteLog;
using DEA2Levels;
using DEAHelper1Leve;
using ReadSettings;
using CreateMetadataFile; // Might need to use this later so leaving it.
using System.Diagnostics;

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
                        await GraphHelper1LevelClass.GetEmailsAttacments1Level(graphClient!, Email);
                    }
                    else
                    {
                        // Calls the function to read ATC emails.
                        await GraphHelper2Levels.GetEmailsAttacmentsAccount(graphClient!, Email);
                    }
                }
                catch (Exception ex)
                {
                    WriteLogClass.WriteToLog(1, $"Exception at email loading: {ex.Message}");
                }
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
                var PathFullDownloadFile = Path.Combine(DownloadFolderPath, DownloadFileName);
                var FileNameOnly = Path.GetFileNameWithoutExtension(PathFullDownloadFile);
                var FileExtention = Path.GetExtension(PathFullDownloadFile);
                var FilePathOnly = Path.GetDirectoryName(PathFullDownloadFile);                
                int Count = 1;

                while (System.IO.File.Exists(PathFullDownloadFile)) // If file exists starts to rename from next file.
                {
                    var NewFileName = string.Format("{0}({1})", FileNameOnly, Count++); // Makes the new file name.
                    PathFullDownloadFile = Path.Combine(FilePathOnly!, NewFileName + FileExtention); // Set tthe new path as the download file path.
                }

                // Full path for the attachment to be downloaded with the attachment name                
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
                var SourceParent = System.IO.Directory.GetParent(SourceFolderPath);
                var SourceFolders = System.IO.Directory.GetDirectories(SourceParent!.FullName, "*.*", SearchOption.AllDirectories);

                foreach (var SourceFolder in SourceFolders)
                {
                    var SourceLastFolder = SourceFolder.Split(Path.DirectorySeparatorChar).Last(); // Get the last folder from the source path.
                    var SourceFiles = System.IO.Directory.GetFiles(SourceFolder, "*.*", SearchOption.AllDirectories); // Get the source file list.
                    var FullDestinationPath = Path.Combine(DestiFolderPath, SourceLastFolder); // Makes the destiantion path with the last folder name
                    
                    if (!System.IO.Directory.Exists(FullDestinationPath)) // Create the folder if not exists.
                    {
                        System.IO.Directory.CreateDirectory(FullDestinationPath);
                    }

                    foreach (var SourceFile in SourceFiles) // Loop throug the files list.
                    {
                        var SourceFileName = System.IO.Path.GetFileName(SourceFile); // Get the source file name.
                        var SourcePath = Path.Combine(SourceFolder, SourceFileName); // Makes the full source path.
                        var DestinationPath = Path.Combine(FullDestinationPath, SourceFileName); // Makes the full destination path.

                        if (!System.IO.Directory.Exists(DestinationPath))
                        {
                            System.IO.File.Move(SourcePath, DestinationPath); // Moves the files to the destination path.

                            WriteLogClass.WriteToLog(3, $"Moving file {SourceFileName}");
                        }
                        else
                        {
                            WriteLogClass.WriteToLog(3, $"File already exists .... skipping to the next file.");
                            continue;
                        }
                    }

                    if (System.IO.Directory.Exists(SourceFolder)) // Delete the file from source if exists.
                    {
                        System.IO.Directory.Delete(SourceFolder, true);
                    }
                }
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
                        MailComment = "Hi,<br /><b>Please do not reply to this email</b><br />. Below email doesn't contain any attachment."; // Can be change with html.
                    }
                    else
                    {
                        MailComment = "Hi,<br /><b>Please do not reply to this email</b><br />Below email's attachment file type is not accepted. Please send attachments as .pdf or .jpg.";
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
                        MailComment = "Hi,<br /><b>Please do not reply to this email</b><br />. Below email doesn't contain any attachment."; // Can be change with html.
                    }
                    else
                    {
                        MailComment = "Hi,<br /><b>Please do not reply to this email</b><br />Below email's attachment file type is not accepted. Please send attachments as .pdf or .jpg.";
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
                if (string.IsNullOrEmpty(ThirdFolderId) && string.IsNullOrEmpty(SecondFolderId))
                {
                    //Graph api call to move the email message.
                    await graphClient!.Users[$"{_Email}"].MailFolders["Inbox"]
                        .ChildFolders[$"{FirstFolderId}"]
                        .Messages[$"{MsgId}"]
                        .Move(DestiId)
                        .Request()
                        .PostAsync();
                }
                else if (string.IsNullOrEmpty(ThirdFolderId))
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
                else
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