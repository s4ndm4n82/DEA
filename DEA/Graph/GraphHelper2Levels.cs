using Microsoft.Graph;
using System.Diagnostics.CodeAnalysis;
using DEA;
using System.Linq;

namespace DEA2Levels
{
    internal class GraphHelper2Levels
    {
        public static async Task GetEmailsAttacmentsAccount([NotNull] GraphServiceClient graphClient, string _Email)
        {
            var ImportMainFolder = @"D:\Import\"; //Path to import folder. 

            var StaticThirdSubFolderID = string.Empty;

            var MailFilterString = "isRead eq false";

            var DateToDay = DateTime.Now.ToString("dd.MM.yyyy");

            // TODO: 1. Enable the bwlo two queries before server testing.
            var SearchOption = new List<QueryOption>
            {
                //new QueryOption("search", $"%22hasAttachments:true received:{DateToDay}%22")
                //new QueryOption("search", $"%22hasAttachments:true%22") //remove this.
            };

            try
            {
                /*var DataConnector = new Microsoft.Graph.ExternalConnectors.ExternalConnection
                {
                    Id = "DEADaemon",
                    Name = "DEA Daemon",
                    Description = "DEA downloader for mail attachments."
                };

                try
                {
                    await graphClient.External.Connections
                    .Request()
                    .AddAsync(DataConnector);

                    Console.WriteLine("Connection Sue");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Excption Thrown: {0}", ex.Message);
                }*/
                

                //Top level of mail boxes like user inbox.
                var FirstSubFolderIDs = await graphClient.Users[$"{_Email}"].MailFolders["Inbox"].ChildFolders                    
                    .Request()
                    .Select(fid => new
                    {
                        fid.Id,
                        fid.DisplayName
                    })
                    .GetAsync();

                    foreach (var FirstSubFolderID in FirstSubFolderIDs)
                    {
                        if (FirstSubFolderID.Id != null)
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
                            if (SecondSubFolderID.DisplayName == "Processing")
                            {
                                Console.WriteLine("Second level folder: {0}", SecondSubFolderID.DisplayName);
                                // Third level of subfolders under the inbox.
                                var GetMessageAttachments = await graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                                    .ChildFolders[$"{FirstSubFolderID.Id}"]
                                    .ChildFolders[$"{SecondSubFolderID.Id}"]
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
                                    .Top(4) //Increase this to 40                                    
                                    .GetAsync();

                                // Counts the with attachments.
                                var MessageCount = GetMessageAttachments.Count();
                                Console.WriteLine("Message Count: {0}", MessageCount);
                                if (MessageCount != 0)
                                {
                                    // Looping through the messages.
                                    foreach (var Message in GetMessageAttachments)
                                    {
                                        // Check if the mail has an attachment or not.
                                       var HasFileAttched = Message.HasAttachments;

                                        Console.WriteLine("Attachment: {0}", HasFileAttched);

                                        // If the file attached variable is true then the download will start.
                                        if (HasFileAttched == true)
                                        {
                                            //Console.WriteLine("{0} Email Subject: {1}", _Email, Message.Subject);
                                            
                                            // Counting the aount of messages with attachments. To loop through below.
                                            var AttachmentCount = Message.Attachments.Count;

                                            // Assigning display names.                                            
                                            var FirstFolderName = FirstSubFolderID.DisplayName;
                                            //var SecondFolderName = SecondSubFolderID.DisplayName;

                                            // Creating the destnation folders.
                                            string[] MakeDestinationFolderPath = { ImportMainFolder, _Email, FirstFolderName };
                                            var DestinationFolderPath = Path.Combine(MakeDestinationFolderPath);

                                            // FolderNameRnd creates a 10 digit folder name. CheckFolder returns the download path.
                                            // This has to be called here. Don't put it within the for loop or it will start calling this
                                            // each time and make folder for each file.
                                            //var PathFullDownloadFolder = Path.Combine(GraphHelper.CheckFolders(), GraphHelper.FolderNameRnd(10));
                                            //Console.WriteLine("Path Download Folder: {0}", PathFullDownloadFolder);
                                            //Console.WriteLine("Path Destination Filder: {0}", DestinationFolderPath);

                                            //TODO: Found how to bypass the issue of empty folder. Refine the solution below. Below IF checks wether theres any
                                            //name element containing the accepted extention if it's there then rest of the code withh run if not it will be skipped.
                                            string[] EndExtention = { ".pdf" };

                                            string PathFullDownloadFolder = string.Empty;

                                            for (int i = 0; i < EndExtention.Length; ++i)
                                            {
                                                var TestExt = Message.Attachments.Where(x => x.Name.ToLower().EndsWith(EndExtention[i]));

                                                if (TestExt.Any(y => y.Name.ToLower().Contains(EndExtention[i])))
                                                {
                                                    Console.WriteLine("{0} Email Subject: {1}", _Email, Message.Subject);
                                                    PathFullDownloadFolder = Path.Combine(GraphHelper.CheckFolders(), GraphHelper.FolderNameRnd(10));
                                                    Console.WriteLine("Path Download Folder in side if: {0}", PathFullDownloadFolder);                                                    
                                                    Console.WriteLine("Esctention: {0}", EndExtention);
                                                    foreach (var Attachment in TestExt)
                                                    {
                                                        Console.WriteLine("Attachment Name: {0}", Attachment.Name);

                                                        var AttachedItem = (FileAttachment)Attachment;// Attachment properties.
                                                        string AttachedItemName = AttachedItem.Name;// Attachment name.
                                                        byte[] AttachedItemBytes = AttachedItem.ContentBytes;// Attachment bytes to be saved on to the disk.
                                                       //PathFullDownloadFolder = Path.Combine(GraphHelper.CheckFolders(), GraphHelper.FolderNameRnd(10));
                                                       
                                                        // Download the files and saves them on to the drive.
                                                        GraphHelper.DownloadAttachedFiles(PathFullDownloadFolder, AttachedItemName, AttachedItemBytes);
                                                    }
                                                    Console.WriteLine(Environment.NewLine);
                                                    string[] DownloadFileExistTest = System.IO.Directory.GetFiles(PathFullDownloadFolder);
                                                }
                                                else
                                                {
                                                    Console.WriteLine(Environment.NewLine);
                                                    Console.WriteLine("No maching attachment");
                                                }
                                                
                                                /*foreach (var Attachment in TestExt!)
                                                {
                                                    var AttachedItem = (FileAttachment)Attachment;// Attachment properties.
                                                    string AttachedItemName = AttachedItem.Name;// Attachment name.
                                                    byte[] AttachedItemBytes = AttachedItem.ContentBytes;// Attachment bytes to be saved on to the disk.
                                                    //PathFullDownloadFolder = Path.Combine(GraphHelper.CheckFolders(), GraphHelper.FolderNameRnd(10));
                                                    // Download the files and saves them on to the drive.
                                                    GraphHelper.DownloadAttachedFiles(PathFullDownloadFolder, AttachedItemName, AttachedItemBytes);
                                                }*/

                                                //string[] DownloadFileExistTest = System.IO.Directory.GetFiles(PathFullDownloadFolder);
                                            }
                                            //var TestExt = Message.Attachments.Where(x => x.Name.ToLower().EndsWith(".pdf") || x.Name.ToLower().EndsWith(".jpg")); 
                                            //var TestExt = Message.Attachments.Where(x => x.Name.ToLower().EndsWith(".pdf"));
                                            //Console.WriteLine("Path Download Folder: {0}", PathFullDownloadFolder);

                                            // Loops through the attachments with in a single email.
                                            /*for (int i = 0; i < AttachmentCount; ++i)
                                            {
                                                // Get the message according to index.
                                                var Attachment = Message.Attachments[i];

                                                // Get the attachment extention only to check if it accepted or not.
                                                var AttachmentExtention = Path.GetExtension(Attachment.Name).Replace(".","").ToLower();
                                                
                                                if (AttachmentExtention == "pdf")// Check the attachment extention.
                                                {
                                                    var AttachedItem = (FileAttachment)Attachment;// Attachment properties.
                                                    string AttachedItemName = AttachedItem.Name;// Attachment name.
                                                    byte[] AttachedItemBytes = AttachedItem.ContentBytes;// Attachment bytes to be saved on to the disk.

                                                    // Download the files and saves them on to the drive.
                                                    GraphHelper.DownloadAttachedFiles(PathFullDownloadFolder, AttachedItemName, AttachedItemBytes);
                                                }
                                            }*/
                                            // Checking the folder and files with in it exsists.
                                            //string[] DownloadFolderExistTest = System.IO.Directory.GetDirectories(GraphHelper.CheckFolders()); // Use the main path not the entire download path
                                            //string[] DownloadFileExistTest = System.IO.Directory.GetFiles(PathFullDownloadFolder);

                                            // Checking if the folders are empty or not.
                                            /*if (DownloadFolderExistTest.Length != 0 && DownloadFileExistTest.Length != 0)
                                            {
                                                // Moves the downloaded files to destination folder. This would create the folder path if it's missing.
                                                if (GraphHelper.MoveFolder(PathFullDownloadFolder, DestinationFolderPath))
                                                {
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
                                                            //.ChildFolders[$"{SecondSubFolderID.Id}"]
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
                                                                if (await GraphHelper.MoveEmails(FirstSubFolderID.Id, SecondSubFolderID.Id, StaticThirdSubFolderID, MessageID, MoveDestinationID, _Email))
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
                                            }*/
                                            Console.WriteLine("-------------------------------------------");
                                        }
                                        else
                                        {
                                            if (HasFileAttched == false)
                                            {
                                                // Search for the subfolder named error.
                                                var FolderSearchOption2 = new List<QueryOption>
                                                        {
                                                            new QueryOption ("filer", $"displayName eq %27Error%27")
                                                        };

                                                //Loop thorugh to Select only error folder from the subfolders.
                                                var ErroFolderDetails = await graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                                                    .ChildFolders[$"{FirstSubFolderID.Id}"]
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
                                                        var StatiicSecondFolderID = string.Empty;

                                                        // Email is beeing forwarded.
                                                        var ForwardDone = await GraphHelper.ForwardEmtpyEmail(FirstSubFolderID.Id, StatiicSecondFolderID, ErrorFolderId, MessageID2, _Email);

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
                                                        if (await GraphHelper.MoveEmails(FirstSubFolderID.Id, SecondSubFolderID.Id, StaticThirdSubFolderID, MessageID2, ErrorFolderId, _Email))
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
                                        Console.WriteLine(Environment.NewLine);
                                    }
                                }
                            }
                            continue;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception at end of main foreach: {0}", ex.Message);
            }
            
        }
    }
}
