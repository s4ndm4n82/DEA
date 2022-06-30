using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System.Text.RegularExpressions;
using DEA2Levels;

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
                Console.WriteLine("Exception: {0}", ex.Message);
                return Task.FromResult(false);
            }
        }

        public static async Task<string> GetAuthTokenWithOutUser(string ClientID, string InstanceID, string TenantID, string ClientSecret, string[] scopes)
        {
            string Authority = string.Concat(InstanceID, TenantID);

            Application = ConfidentialClientApplicationBuilder.Create(ClientID)
                          .WithClientSecret(ClientSecret)
                          .WithAuthority(new Uri(Authority))
                          .Build();
            try
            {
                AuthToken = await Application.AcquireTokenForClient(scopes).ExecuteAsync();                
            }
            catch (MsalUiRequiredException ex)
            {
                // The application doesn't have sufficient permissions.
                // - Did you declare enough app permissions during app creation?
                // - Did the tenant admin grant permissions to the application?
                Console.WriteLine("Exception: {0}", ex.Message);
            }
            catch (MsalServiceException ex) when (ex.Message.Contains("AADSTS70011"))
            {
                // Invalid scope. The scope has to be in the form "https://resourceurl/.default"
                // Mitigation: Change the scope to be as expected.
                Console.WriteLine("Scope provided is not supported");
            }

            
            return AuthToken!.AccessToken;

        }

        public static async Task InitializGetAttachment()
        {
            List<string> EmailCheckList = new List<string>();

            string[] EmailsList =
            {
                "accounting@efakturamottak.no",
                /*"accounting02@efakturamottak.no",
                "accounting03@efakturamottak.no",
                "accounting04@efakturamottak.no",
                "accounting05@efakturamottak.no",
                "atc@efakturamottak.no",
                "atc02@efakturamottak.no"*/
            };

            EmailCheckList.AddRange(EmailsList);

            foreach (string Email in EmailCheckList)
            {
                //TODO: 1. Accounting emails folder structer ends at level 2 subfolders. Seems I've to create another function for those but in a different class.

                try
                {
                    Regex EmailRegEx = new Regex(@"^accounting+(?=[0-9]{0,3}@[a-z]+[\.][a-z]{2,3})");
                    if (EmailRegEx.IsMatch(Email))
                    {
                        Console.WriteLine("Accessing email {0}", Email);
                        await GraphHelper2Levels.GetEmailsAttacmentsAccount(graphClient!, Email);
                        Console.WriteLine(Environment.NewLine);
                    }
                    else
                    {
                        Console.WriteLine("Executing await GetAttachmentTodayAsync(Email) for {0}", Email);
                        //await GetAttachmentTodayAsync(Email);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Exception thrown: {0}", ex.Message);
                }
            }
        }

        public static async Task GetAttachmentTodayAsync(string _Email)
        {
            var ImportMainFolder = @"D:\Import\"; //Path to import folder. 

            var DateToDay = DateTime.Now.ToString("dd.MM.yyyy");

            // TODO: 1. Enable the bwlo two queries before server testing.
            var SearchOptions = new List<QueryOption>
            {
                //new QueryOption("search", $"%22hasAttachments:true received:{DateToDay}%22")
                //new QueryOption("search", $"%22hasAttachments:true%22") //remove this.
            };

            try
            {
                /* Can't use .Me with application permissions. It only can be used with delegated permissions.
                 var FirstSubFolderIDs = await graphClient.Me.MailFolders["Inbox"].ChildFolders */

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
                        Console.WriteLine("First level folder: {0}", FirstSubFolderID.DisplayName);
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

                        foreach (var SecondSubFolderID in SecondSubFolderIDs)
                        {
                            if(SecondSubFolderID.Id != null)
                            {
                                Console.WriteLine("Second level folder: {0}", SecondSubFolderID.DisplayName);
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
                                    Console.WriteLine("Third level folder: {0}", ThirdSubFolderID.DisplayName);
                                    if (ThirdSubFolderID.DisplayName == "Processing")
                                    {                                        
                                        // Looping through the emails in the subfolder "New".
                                        var GetMessageAttachments = await graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                                            .ChildFolders[$"{FirstSubFolderID.Id}"]
                                            .ChildFolders[$"{SecondSubFolderID.Id}"]
                                            .ChildFolders[$"{ThirdSubFolderID.Id}"]
                                            .Messages
                                            //.Request(SearchOption) // Uncomment this before thesting.
                                            .Request()
                                            .Expand("attachments")
                                            .Select(gma => new
                                            {   
                                                gma.Id,
                                                gma.Subject,
                                                gma.HasAttachments,                                               
                                                gma.Attachments
                                            })
                                            .GetAsync();

                                        // Gets the message count.
                                        var MessageCount = GetMessageAttachments.Count;

                                        if (MessageCount != 0)
                                        {
                                            Console.WriteLine("Messages");

                                            // Loops through the emails.
                                            foreach (var Message in GetMessageAttachments)
                                            {   
                                                Console.WriteLine("Subjec: {0}", Message.Subject);

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
                                                    string[] MakeDestinationFolderPath = { ImportMainFolder, FirstFolderName, SecondFolderName };
                                                    var DestinationFolderPath = Path.Combine(MakeDestinationFolderPath);

                                                    // FolderNameRnd creates a 10 digit folder name. CheckFolder returns the download path.
                                                    // This has to be called here. Don't put it within the for loop or it will start calling this
                                                    // each time and make folder for each file.
                                                    var PathFullDownloadFolder = Path.Combine(CheckFolders(), FolderNameRnd(10));

                                                    // Loops through the attachments with in a single email.
                                                    for (int i = 0; i < AttachmentCount; ++i)
                                                    {
                                                        // Get the message according to index.
                                                        var Attachment = Message.Attachments[i];

                                                        // Get the attachment extention only to check if it accepted or not.
                                                        var AttachmentExtention = Path.GetExtension(Attachment.Name).Replace(".", "").ToLower();

                                                        if (AttachmentExtention == "pdf")// Check the attachment extention.
                                                        {
                                                            var AttachedItem = (FileAttachment)Attachment;// Attachment properties.
                                                            string AttachedItemName = AttachedItem.Name;// Attachment name.
                                                            byte[] AttachedItemBytes = AttachedItem.ContentBytes;// Attachment bytes to be saved on to the disk.

                                                            // Download the files and saves them on to the drive.
                                                            DownloadAttachedFiles(PathFullDownloadFolder, AttachedItemName, AttachedItemBytes);
                                                        }
                                                    }
                                                    // Checking the folder and files with in it exsists.
                                                    string[] DownloadFolderExistTest = System.IO.Directory.GetDirectories(CheckFolders()); // Use the main path not the entire download path
                                                    string[] DownloadFileExistTest = System.IO.Directory.GetFiles(PathFullDownloadFolder);

                                                    // Checking if the folders are empty or not.
                                                    if (DownloadFolderExistTest.Length != 0 && DownloadFileExistTest.Length != 0)
                                                    {
                                                        // Moves the downloaded files to destination folder. This would create the folder path if it's missing.
                                                        if (MoveFolder(PathFullDownloadFolder, DestinationFolderPath))
                                                        {
                                                            // Search option sets the $filter query to only get the folder named downloaded.
                                                            var FolderSearchOption = new List<QueryOption>
                                                        {
                                                            new QueryOption ("filter", $"displayName eq %27Exported%27")
                                                        };

                                                            try
                                                            {
                                                                // Loop through and selects only the downloaded folder.
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
                                                                        if (await MoveEmails(FirstSubFolderID.Id, SecondSubFolderID.Id, ThirdSubFolderID.Id, MessageID, MoveDestinationID, _Email))
                                                                        {
                                                                            Console.WriteLine("Email Moved ....");
                                                                        }
                                                                        else
                                                                        {
                                                                            Console.WriteLine("Email Did not Move ....");
                                                                        }                                                                        
                                                                    }
                                                                }
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                Console.WriteLine($"Exception Thrown: {ex.Message}");
                                                            }

                                                        }
                                                    }
                                                    Console.WriteLine("-------------------------------------------");
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
                                                                    Console.WriteLine($"Email Forwarded to {ForwardDone.Item1}");
                                                                }
                                                                else
                                                                {
                                                                    Console.WriteLine($"Email not Forawarded. Exception: {ForwardDone.Item1}");
                                                                }

                                                                // Moves the empty emails to error folder once forwarding is done.
                                                                if (await MoveEmails(FirstSubFolderID.Id, SecondSubFolderID.Id, ThirdSubFolderID.Id, MessageID2, ErrorFolderId, _Email))
                                                                {
                                                                    Console.WriteLine($"Mail Moved to {ErrorFolder.DisplayName} Folder ....\n");                                                                  
                                                                }
                                                                else
                                                                {
                                                                    Console.WriteLine($"Mail was Not Moved to {ErrorFolder.DisplayName} Folder ....\n");
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            Console.WriteLine("\n");                                            
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
                Console.WriteLine($"Error: {ex.Message}");
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
        public static string CheckFolders()
        {
            // Get current execution path.
            string? PathRootFolder = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string DownloadFolderName = "Download";
            string PathDownloadFolder = Path.Combine(PathRootFolder!, DownloadFolderName);

            // Check if download folder exists. If not creates the fodler.
            if (!System.IO.Directory.Exists(PathDownloadFolder))
            {
                try
                {
                    System.IO.Directory.CreateDirectory(PathDownloadFolder);
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Error getting events: {ex.Message}");
                }
            }

            return PathDownloadFolder;
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
                    Console.WriteLine($"Exceptio at download folder creation: {ex.Message}");
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
                Console.WriteLine($"Exception at download path: {ex.Message}");
                return false;
            }
        }

        // Move the folder to main import folder on the loca machine.
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
                Console.WriteLine($"Error getting event: {ex.Message}");
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
                Console.WriteLine($"Exception Thrown: {ex.Message}");

                return false;
            }
        }
    }
}