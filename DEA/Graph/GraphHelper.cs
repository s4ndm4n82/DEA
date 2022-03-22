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
                                                Console.WriteLine("Subjec: {0}", Message.Subject);
                                                Console.WriteLine("Has Attachment: {0}", Message.HasAttachments);

                                                foreach (var Attachment in Message.Attachments)
                                                {
                                                    Console.WriteLine("Folder 1 Name: {0}\n", FirstSubFolderID.DisplayName);
                                                    Console.WriteLine("Folder 2 Name: {0}\n", SecondSubFolderID.DisplayName);
                                                    Console.WriteLine("Folder 3 Name: {0}\n", ThirdSubFolderID.DisplayName);
                                                    Console.WriteLine("\nRandom Number: {0}", FolderNameRnd(10));
                                                    Console.WriteLine("\nAttachment: {0}", Attachment.Name);
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
    }
}