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
            string PathImportFolder = @"D:\Import\"; //Path to import folder.            

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

                                        var MessageCount = GetMessageAttachments.Count();

                                        if (MessageCount != 0)
                                        {
                                            Console.WriteLine("Messages");
                                            foreach (var Message in GetMessageAttachments)
                                            {
                                                Console.WriteLine("\n");
                                                Console.WriteLine("Subjec: {0}", Message.Subject);
                                                Console.WriteLine("Has Attachment: {0}", Message.HasAttachments);
                                                

                                                var AttachmentCount = Message.Attachments.Count();

                                                Console.WriteLine("Attachment Count: {0}", AttachmentCount);

                                                for (int i = 0; i < AttachmentCount; i++)
                                                {
                                                    var Attachment = Message.Attachments[i];
                                                    Console.WriteLine("Attachment Name: {0}", Attachment.Name);                                                    
                                                }
                                                
                                                /*foreach (var Attachment in Message.Attachments)
                                                {
                                                    Console.WriteLine("Folder 1 Name: {0}\n", FirstSubFolderID.DisplayName);
                                                    Console.WriteLine("Folder 2 Name: {0}\n", SecondSubFolderID.DisplayName);
                                                    Console.WriteLine("Folder 3 Name: {0}\n", ThirdSubFolderID.DisplayName);
                                                    Console.WriteLine("\nAttachment: {0}", Attachment.Name);

                                                    var AttachedItem = (FileAttachment)Attachment;//Attachment properties.

                                                    //FolderNameRnd creates a 10 digit folder name. CheckFolder returns the download path.
                                                    var PathFullDownloadFolder = Path.Combine(CheckFolders(), FolderNameRnd(10));
                                                    
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
                                                    System.IO.File.WriteAllBytes(PathFullDownloadFile, AttachedItem.ContentBytes);*/

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
                                                    }
                                                }*/
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
    }
}