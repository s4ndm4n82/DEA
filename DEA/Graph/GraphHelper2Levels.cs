using Microsoft.Graph;
using System.Diagnostics.CodeAnalysis;
using Serilog;
using DEA;
using ReadSettings;
using WriteLog;

namespace DEA2Levels
{
    internal class GraphHelper2Levels
    {
        public static async Task GetEmailsAttacmentsAccount([NotNull] GraphServiceClient graphClient, string _Email)
        {
            // Parameters read from the config files.
            var ConfigParam = new ReadSettingsClass();

            bool DateSwitch = ConfigParam.DateFilter;
            int MaxAmountOfEmails = ConfigParam.MaxLoadEmails;
            string ImportFolderPath = Path.Combine(ConfigParam.ImportFolderLetter, ConfigParam.ImportFolderPath);

            //var ImportFolderPath = @"D:\Import\"; //Path to import folder. 

            var StaticThirdSubFolderID = string.Empty; // Variable to null the third folder ID variable.

            var DateToDay = DateTime.Now.ToString("dd.MM.yyyy");

            // TODO: 1. Enable one out query from below before server testing.
            List<QueryOption> SearchOptions;

            if (DateSwitch)
            {
                SearchOptions = new List<QueryOption>
                {
                    new QueryOption("search", $"%22received:{DateToDay}%22")
                    //new QueryOption("search", $"%22hasAttachments:true received:{DateToDay}%22")
                    //new QueryOption("search", $"%22hasAttachments:true%22") //remove this.
                 };
            }
            else
            {
                SearchOptions = new List<QueryOption> { };
            }

            try
            {
                /* Creating a new connection. Need to use this later
                var DataConnector = new Microsoft.Graph.ExternalConnectors.ExternalConnection
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

                WriteLogClass.WriteToLog(3, $"Processing Email: {_Email}");

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
                                WriteLogClass.WriteToLog(3, $"Processing folder path {FirstSubFolderID.DisplayName}->{SecondSubFolderID.DisplayName}");
                                // Third level of subfolders under the inbox.
                                var GetMessageAttachments = await graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                                    .ChildFolders[$"{FirstSubFolderID.Id}"]
                                    .ChildFolders[$"{SecondSubFolderID.Id}"]
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
                                    .Top(MaxAmountOfEmails) // Increase this to 40                                    
                                    .GetAsync();

                                // Counts the with attachments.
                                var MessageCount = GetMessageAttachments.Count();

                                if (MessageCount != 0)
                                {
                                    // Looping through the messages.
                                    foreach (var Message in GetMessageAttachments)
                                    {
                                        // Check if the mail has an attachment or not.
                                        var HasFileAttched = Message.HasAttachments;

                                        // If the file attached variable is true then the download will start.
                                        if (HasFileAttched == true)
                                        {
                                            // Counting the aount of messages with attachments. To loop through below.
                                            var AttachmentCount = Message.Attachments.Count;

                                            // Assigning display names.                                            
                                            var FirstFolderName = FirstSubFolderID.DisplayName;

                                            // Creating the destnation folders.
                                            string[] MakeDestinationFolderPath = { ImportFolderPath, _Email, FirstFolderName };
                                            var DestinationFolderPath = Path.Combine(MakeDestinationFolderPath);

                                            // Variable used to store all the accepted extentions.
                                            string[] AcceptedExtention = ConfigParam.AllowedExtentions;

                                            // Initilizing the download folder path variable.
                                            string PathFullDownloadFolder = string.Empty;

                                            // Used to set the status of if the email is moved or not.
                                            // If moved this will be true. If not the code for forwarding will be executed.
                                            var EmailMoveStatus = false;

                                            // For loop to go through all the extentions from extentions variable.
                                            for (int i = 0; i < AcceptedExtention.Length; ++i)
                                            {
                                                var AcceptedExtensionCollection = Message.Attachments.Where(x => x.Name.ToLower().EndsWith(AcceptedExtention[i]));

                                                if (AcceptedExtensionCollection.Any(y => y.Name.ToLower().Contains(AcceptedExtention[i])))
                                                {
                                                    WriteLogClass.WriteToLog(3, "Collection check succeeded ...");

                                                    // FolderNameRnd creates a 10 digit folder name. CheckFolder returns the download path.
                                                    // This has to be called here. Don't put it within the for loop or it will start calling this
                                                    // each time and make folder for each file. Also calling this out side of the extentions FOR loop.
                                                    // causes an exception error at the "DownloadFileExistTest" test due file not been available.
                                                    PathFullDownloadFolder = Path.Combine(GraphHelper.CheckFolders("Download"), GraphHelper.FolderNameRnd(10));


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
                                                            // Should mark and download the itemattacment which is the correct attachment.
                                                            var TrueAttachment = await graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                                                                            .ChildFolders[$"{FirstSubFolderID.Id}"]
                                                                            .ChildFolders[$"{SecondSubFolderID.Id}"]
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
                                                            GraphHelper.DownloadAttachedFiles(PathFullDownloadFolder, TrueAttachmentName, TruAttachmentBytes);

                                                            WriteLogClass.WriteToLog(3, $"Downloaded attachments from {Message.Subject}");

                                                            // Directory and file existence check. If not exists it will not return anything.
                                                            string[] DownloadFolderExistTest = System.IO.Directory.GetDirectories(GraphHelper.CheckFolders("Download")); // Use the main path not the entire download path
                                                            string[] DownloadFileExistTest = System.IO.Directory.GetFiles(PathFullDownloadFolder); // This causs an erro when the file is not there.

                                                            if (DownloadFolderExistTest.Length != 0 && DownloadFileExistTest.Length != 0)
                                                            {
                                                                WriteLogClass.WriteToLog(3, "Moving downloaded files to local folder.");
                                                                // Moves the downloaded files to destination folder. This would create the folder path if it's missing.
                                                                if (GraphHelper.MoveFolder(PathFullDownloadFolder, DestinationFolderPath))
                                                                {
                                                                    WriteLogClass.WriteToLog(3, "File moved successfully.");
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
                                                                                    WriteLogClass.WriteToLog(3, $"Email {Message.Subject} moved to export folder ...");
                                                                                    EmailMoveStatus = true;
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
                                                                        WriteLogClass.WriteToLog(1, $"Exception at attachment download area 2level: {ex.Message}");
                                                                    }

                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (!EmailMoveStatus)
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
                                                                var AttachmentStatus = 1; // Extension not accepted.
                                                                var StatiicSecondFolderID = string.Empty;

                                                                // Email is beeing forwarded.
                                                                var ForwardDone = await GraphHelper.ForwardEmtpyEmail(FirstSubFolderID.Id, StatiicSecondFolderID, ErrorFolderId, MessageID2, _Email, AttachmentStatus);

                                                                // After forwarding checks if the action returned true.
                                                                // Item2 is the bool value returned.
                                                                // Item1 is the maile address.
                                                                if (ForwardDone.Item2)
                                                                {
                                                                    WriteLogClass.WriteToLog(3, $"Email forwarded to {ForwardDone.Item1}");
                                                                }
                                                                else
                                                                {
                                                                    WriteLogClass.WriteToLog(3, $"Email not forwarded to {ForwardDone.Item1}");
                                                                }

                                                                // Moves the empty emails to error folder once forwarding is done.
                                                                if (await GraphHelper.MoveEmails(FirstSubFolderID.Id, SecondSubFolderID.Id, StaticThirdSubFolderID, MessageID2, ErrorFolderId, _Email))
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
                                            if (HasFileAttched == false)
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
                                                            var AttachmentStatus = 0; // No Attachment
                                                            var StatiicSecondFolderID = string.Empty;

                                                            // Email is beeing forwarded.
                                                            var ForwardDone = await GraphHelper.ForwardEmtpyEmail(FirstSubFolderID.Id, StatiicSecondFolderID, ErrorFolderId, MessageID2, _Email, AttachmentStatus);

                                                            // After forwarding checks if the action returned true.
                                                            // Item2 is the bool value returned.
                                                            // Item1 is the maile address.
                                                            if (ForwardDone.Item2)
                                                            {
                                                                WriteLogClass.WriteToLog(3, $"Email forwarded to {ForwardDone.Item1}");
                                                            }
                                                            else
                                                            {
                                                                WriteLogClass.WriteToLog(3, $"Email not forwarded to {ForwardDone.Item1}");
                                                            }

                                                            // Moves the empty emails to error folder once forwarding is done.
                                                            if (await GraphHelper.MoveEmails(FirstSubFolderID.Id, SecondSubFolderID.Id, StaticThirdSubFolderID, MessageID2, ErrorFolderId, _Email))
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
                                                    WriteLogClass.WriteToLog(1, $"Exception at error folder mover 2level: {ex.Message}");
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
                WriteLogClass.WriteToLog(1, $"Exception at end of main foreach 2level: {ex.Message}");
            }
        }
    }
}
