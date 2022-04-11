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
                new QueryOption("search", $"%22hasAttachments:true%22") //remove this.
            };

            try
            {
                //TODO: 1. Save the downloded files into a local folder called downloads with in the program folder.
                //TODO: 2. Create a folder with uniq name to save the downloaded attachment files.
                //TODO: 3. Make a custom path to move the downloaded files.
                //TODO: 4. Check for folders before moving the files.
                //TODO: 5. Once done, move it to "Import" folder according to the created custom path.

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

                                foreach(var ThirdSubFolderID in ThirdSubFolderIDs)
                                {
                                    
                                    if(ThirdSubFolderID.Id != null)
                                    {
                                        var GetMessageAttachments = await graphClient.Me.MailFolders["Inbox"]
                                            .ChildFolders[$"{FirstSubFolderID.Id}"]
                                            .ChildFolders[$"{SecondSubFolderID.Id}"]
                                            .ChildFolders[$"{ThirdSubFolderID.Id}"]
                                            .Messages
                                            .Request(SearchOption)
                                            .Expand("attachments")
                                            .Select(gma => new
                                            {
                                                gma.Subject,
                                                gma.HasAttachments,
                                                gma.ConversationIndex,                                                
                                                gma.Attachments
                                            })
                                            .GetAsync();

                                        //Get Message count that includes attachments
                                        var MessageCount = GetMessageAttachments.Count();

                                        if (MessageCount != 0)
                                        {
                                            Console.WriteLine("Messages");

                                            foreach (var Message in GetMessageAttachments)
                                            {   
                                                Console.WriteLine("Subjec: {0}", Message.Subject);
                                                //Counting the aount of messages with attachments. To loop through below.
                                                var AttachmentCount = Message.Attachments.Count();

                                                //Assigning display names.
                                                var FirstFolderName = FirstSubFolderID.DisplayName;
                                                var SecondFolderName = SecondSubFolderID.DisplayName;                                                

                                                //Creating the destnation folders.
                                                string[] MakeDestinationFolderPath = { ImportMainFolder, FirstFolderName, SecondFolderName };
                                                var DestinationFolderPath = Path.Combine(MakeDestinationFolderPath);                                                

                                                //FolderNameRnd creates a 10 digit folder name. CheckFolder returns the download path.
                                                //This has to be called here. Don't put it with in the for loop or it will start calling this
                                                //each time and make folder for each file.
                                                var PathFullDownloadFolder = Path.Combine(CheckFolders(), FolderNameRnd(10));

                                                //TODO: Was looking at while loops to call folder creation once.
                                                for (int i = 0; i < AttachmentCount; ++i)
                                                {
                                                    //Get the message according to index.
                                                    var Attachment = Message.Attachments[i];                                                    
                                                    var AttachmentExtention = Path.GetExtension(Attachment.Name).Replace(".","").ToLower();
                                                    
                                                    if (AttachmentExtention == "pdf")
                                                    {                                                        
                                                        var AttachedItem = (FileAttachment)Attachment;//Attachment properties.
                                                        if (!System.IO.Directory.Exists(PathFullDownloadFolder))
                                                        {
                                                            try
                                                            {
                                                                System.IO.Directory.CreateDirectory(PathFullDownloadFolder);
                                                            }
                                                            catch (Exception ex)
                                                            {
                                                                Console.WriteLine($"Error getting events: {ex.Message}");
                                                            }
                                                        }
                                                        //Fulle path for the attachment to be downloaded with the attachment name
                                                        var PathFullDownloadFile = Path.Combine(PathFullDownloadFolder, AttachedItem.Name);
                                                        System.IO.File.WriteAllBytes(PathFullDownloadFile, AttachedItem.ContentBytes);

                                                        if (System.IO.File.Exists(PathFullDownloadFile))
                                                        {
                                                            //TODO: Work on the file moving.
                                                            //MoveFolder(PathFullDownloadFolder, DestinationFolderPath);
                                                        }
                                                    }
                                                }
                                                Console.WriteLine("-------------------------------------------");
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
            //await GraphHelper.GetAttachmentTodayAsync();
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

        //Move the folder to main import folder.
        private static bool MoveFolder(string SourceFolderPath, string DestiFolderPath)
        {
            var MoveDone = false;

            try
            {
                if (System.IO.Directory.Exists(SourceFolderPath))
                {
                    System.IO.Directory.Delete(SourceFolderPath, true);
                    System.IO.File.Move(SourceFolderPath, DestiFolderPath);
                    Console.WriteLine("Moving Done.");
                    MoveDone = true;
                }
                else
                {
                    System.IO.File.Move(SourceFolderPath, DestiFolderPath);
                    Console.WriteLine("Moving Done.");
                    MoveDone = true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Error getting event: {ex.Message}");
            }

            return MoveDone;
        }
    }
}