using Azure.Core;
using Azure.Identity;
using Microsoft.Graph;

namespace DEA
{
    public class GraphHelper
    {
        private static DeviceCodeCredential? tokenCredentials;
        private static GraphServiceClient? graphClient;

        public static void Initialize(string clientID, string[] scopes,
                                      Func<DeviceCodeInfo, CancellationToken, Task> callBack)
        {
            tokenCredentials = new DeviceCodeCredential(callBack, clientID);
            graphClient = new GraphServiceClient(tokenCredentials, scopes);
        }

        public static async Task<string> GetAccessTokenAsync(string[] scopes)
        {
            var context = new TokenRequestContext(scopes);
            var response = await tokenCredentials.GetTokenAsync(context);
            return response.Token;
        }

        public static async Task<User> GetMeAsync()
        {
            try
            {
                //GET /me
                return await graphClient.Me
                    .Request()
                    .Select(u => new {
                        u.DisplayName,
                        u.MailboxSettings
                    })
                    .GetAsync();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error getting signed in useer:{0}", ex.Message);
                return null;
            }
        }

        public static async Task GetAttachmentTodayAsync()
        {
            var ImportMainFolder = @"D:\Import\"; //Path to import folder. 

            var DateToDay = DateTime.Now.ToString("dd.MM.yyyy");

            var SearchOption = new List<QueryOption>
            {
                //new QueryOption("search", $"%22hasAttachments:true received:{DateToDay}%22")
                //new QueryOption("search", $"%22hasAttachments:true%22") //remove this.
            };

            try
            {
                var FirstSubFolderIDs = await graphClient.Me.MailFolders["Inbox"].ChildFolders
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
                        var SecondSubFolderIDs = await graphClient.Me.MailFolders["Inbox"]
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
                                var ThirdSubFolderIDs = await graphClient.Me.MailFolders["Inbox"]
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
                                    if(ThirdSubFolderID.DisplayName == "New")
                                    {
                                        var GetMessageAttachments = await graphClient.Me.MailFolders["Inbox"]
                                            .ChildFolders[$"{FirstSubFolderID.Id}"]
                                            .ChildFolders[$"{SecondSubFolderID.Id}"]
                                            .ChildFolders[$"{ThirdSubFolderID.Id}"]
                                            .Messages
                                            //.Request(SearchOption)
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

                                        //Get Message count that includes attachments
                                        var MessageCount = GetMessageAttachments.Count;

                                        if (MessageCount != 0)
                                        {
                                            Console.WriteLine("Messages");

                                            foreach (var Message in GetMessageAttachments)
                                            {   
                                                Console.WriteLine("Subjec: {0}", Message.Subject);

                                                var HasFileAttached = Message.HasAttachments;
                                                
                                                if (HasFileAttached == true)
                                                {
                                                    //Counting the aount of messages with attachments. To loop through below.
                                                    var AttachmentCount = Message.Attachments.Count;

                                                    //Assigning display names.
                                                    var FirstFolderName = FirstSubFolderID.DisplayName;
                                                    var SecondFolderName = SecondSubFolderID.DisplayName;

                                                    //Creating the destnation folders.
                                                    string[] MakeDestinationFolderPath = { ImportMainFolder, FirstFolderName, SecondFolderName };
                                                    var DestinationFolderPath = Path.Combine(MakeDestinationFolderPath);

                                                    //FolderNameRnd creates a 10 digit folder name. CheckFolder returns the download path.
                                                    //This has to be called here. Don't put it within the for loop or it will start calling this
                                                    //each time and make folder for each file.
                                                    var PathFullDownloadFolder = Path.Combine(CheckFolders(), FolderNameRnd(10));

                                                    //Loops through the attachments with in a single email.
                                                    for (int i = 0; i < AttachmentCount; ++i)
                                                    {
                                                        //Get the message according to index.
                                                        var Attachment = Message.Attachments[i];

                                                        //Get the attachment extention only to check if it accepted or not.
                                                        var AttachmentExtention = Path.GetExtension(Attachment.Name).Replace(".", "").ToLower();

                                                        if (AttachmentExtention == "pdf")//Check the attachment extention.
                                                        {
                                                            var AttachedItem = (FileAttachment)Attachment;//Attachment properties.
                                                            string AttachedItemName = AttachedItem.Name;//Attachment name.
                                                            byte[] AttachedItemBytes = AttachedItem.ContentBytes;//Attachment bytes to be saved on to the disk.

                                                            //Download the files and saves them on to the drive.
                                                            DownloadAttachedFiles(PathFullDownloadFolder, AttachedItemName, AttachedItemBytes);
                                                        }
                                                    }
                                                    //Checking the folder and files with in it exsists.
                                                    string[] DownloadFolderExistTest = System.IO.Directory.GetDirectories(CheckFolders()); //Use the main path not the entire download path
                                                    string[] DownloadFileExistTest = System.IO.Directory.GetFiles(PathFullDownloadFolder);

                                                    //Checking if the folders are not empty.
                                                    if (DownloadFolderExistTest.Length != 0 && DownloadFileExistTest.Length != 0)
                                                    {
                                                        //Moves the downloaded files to destination folder. This would create the folder path if it's missing.
                                                        if (MoveFolder(PathFullDownloadFolder, DestinationFolderPath))
                                                        {
                                                            //Serach option sets the $filter query to only get the folders named downloaded.
                                                            var FolderSearchOption = new List<QueryOption>
                                                        {
                                                            new QueryOption ("filter", $"displayName eq %27Downloaded%27")
                                                        };

                                                            try
                                                            {
                                                                //Loop through and selects only the downloaded folder.
                                                                var DestinationDetails = await graphClient.Me.MailFolders["Inbox"]
                                                                    .ChildFolders[$"{FirstSubFolderID.Id}"]
                                                                    .ChildFolders[$"{SecondSubFolderID.Id}"]
                                                                    .ChildFolders
                                                                    .Request(FolderSearchOption)
                                                                    .GetAsync();

                                                                foreach (var Destination in DestinationDetails)
                                                                {
                                                                    if (Destination.DisplayName == "Downloaded")//Just a backup check of the folder name.
                                                                    {
                                                                        var MessageID = Message.Id;
                                                                        var MoveDestinationID = Destination.Id;

                                                                        //Moves the mail to downloaded folder.
                                                                        if (await MoveEmails(FirstSubFolderID.Id, SecondSubFolderID.Id, ThirdSubFolderID.Id, MessageID, MoveDestinationID))
                                                                        {
                                                                            Console.WriteLine("Email Moved ....");
                                                                        }
                                                                        else
                                                                        {
                                                                            Console.WriteLine("Email Did not Move ....");
                                                                        }
                                                                        //TODO: 1. Need to add move emails to error if no attachments or files not supported.
                                                                        //TODO: 2. Change the permission method to auto from the current manual method.
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
                                                        var FolderSearchOption2 = new List<QueryOption>
                                                        {
                                                            new QueryOption ("filer", $"displayName eq %27Error%27")
                                                        };

                                                        //Loop thorugh to Select only error folder from the subfolders.
                                                        var ErroFolderDetails = await graphClient.Me.MailFolders["Inbox"]
                                                            .ChildFolders[$"{FirstSubFolderID.Id}"]
                                                            .ChildFolders[$"{SecondSubFolderID.Id}"]
                                                            .ChildFolders
                                                            .Request(FolderSearchOption2)
                                                            .GetAsync();

                                                        foreach (var ErrorFolder in ErroFolderDetails)
                                                        {
                                                            if (ErrorFolder.DisplayName == "Error")
                                                            {
                                                                var MessageID2 = Message.Id;
                                                                var ErrorFolderId = ErrorFolder.Id;
                                                                //testing the below code to forwar the email.
                                                                var MsgDetails = await graphClient.Me.MailFolders["Inbox"]
                                                                        .ChildFolders[$"{FirstSubFolderID.Id}"]
                                                                        .ChildFolders[$"{SecondSubFolderID.Id}"]
                                                                        .ChildFolders[$"{ErrorFolderId}"]
                                                                        .Messages[$"{MessageID2}"]
                                                                        .Request()
                                                                        .Select(em => new
                                                                        {
                                                                            em.Subject,
                                                                            em.From
                                                                        })
                                                                        .GetAsync();
                                                                //var msg = MsgDetails.From;

                                                                Console.WriteLine("Email Subject: {0}", MsgDetails.Subject);
                                                                Console.WriteLine("From Email: {0}", MsgDetails.From);

                                                                if (await MoveEmails(FirstSubFolderID.Id, SecondSubFolderID.Id, ThirdSubFolderID.Id, MessageID2, ErrorFolderId))
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
            catch (ServiceException ex)
            {
                Console.WriteLine($"Error getting events: {ex.Message}");
            }
        }

        //Generate the random 10 digit number for the folder name.
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

        //Check the exsistance of the download folders.
        public static string CheckFolders()
        {
            //Get current execution path.
            string PathRootFolder = Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().Location);
            string DownloadFolderName = "Download";
            string PathDownloadFolder = Path.Combine(PathRootFolder, DownloadFolderName);

            //Check if download folder exists. If not creates the fodler.
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
                    Console.WriteLine($"Error getting events: {ex.Message}");
                }
            }

            try
            {
                //Fulle path for the attachment to be downloaded with the attachment name
                var PathFullDownloadFile = Path.Combine(DownloadFolderPath, DownloadFileName);
                System.IO.File.WriteAllBytes(PathFullDownloadFile, DownloadFileData);
                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting events: {ex.Message}");
                return false;
            }
        }

        //Move the folder to main import folder.
        private static bool MoveFolder(string SourceFolderPath, string DestiFolderPath)
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

        //Moves the email to Downloded folder.
        private static async Task<bool> MoveEmails(string FirstFolderId, string SecondFolderId, string ThirdFolderId, string MsgId, string DestiId)
        {
            try
            {
                //Graph api call to move the email message.
                await graphClient.Me.MailFolders["Inbox"]
                    .ChildFolders[$"{FirstFolderId}"]
                    .ChildFolders[$"{SecondFolderId}"]
                    .ChildFolders[$"{ThirdFolderId}"]
                    .Messages[$"{MsgId}"]
                    .Move(DestiId)
                    .Request()
                    .PostAsync();

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