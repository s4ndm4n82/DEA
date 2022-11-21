using Microsoft.Graph;
using System.Diagnostics.CodeAnalysis;
using System.Text.RegularExpressions;
using DEA;
using ReadSettings;
using WriteLog;
using FolderCleaner;
using GetRecipientEmail; // Might need it later.
using CreateMetadataFile; // Might need to use this later so leaving it.

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
                 };
            }
            else
            {
                SearchOptions = new List<QueryOption> { };
            }

            try
            {
                WriteLogClass.WriteToLog(3, $"Processing Email {_Email} ....");

                //Top level of mail boxes like user inbox.
                var FirstSubFolderIDs = await graphClient.Users[$"{_Email}"].MailFolders["Inbox"].ChildFolders                    
                    .Request()
                    .Select(fid => new
                    {
                        fid.Id,
                        fid.DisplayName
                    })
                    .Top(MaxAmountOfEmails)
                    .GetAsync();

                foreach (var FirstSubFolderID in FirstSubFolderIDs.Where(ids => !string.IsNullOrWhiteSpace(ids.Id)))
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
                        .Top(MaxAmountOfEmails)
                        .GetAsync();

                    foreach (var SecondSubFolderID in SecondSubFolderIDs.Where(ids2 => !string.IsNullOrWhiteSpace(ids2.Id)))
                    {
                        WriteLogClass.WriteToLog(3, $"Processing folder path {FirstSubFolderID.DisplayName} -> {SecondSubFolderID.DisplayName}");
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
                        // Looping through the messages.
                        foreach (var Message in GetMessageAttachments)
                        {
                            // Assigning display names.                                            
                            var FirstFolderName = FirstSubFolderID.DisplayName;
                            var SecondFolderName = SecondSubFolderID.DisplayName;

                            // Extracted recipient email for creating the folder path.
                            //var RecipientEmail = GetRecipientEmailClass.GetRecipientEmail(graphClient, FirstSubFolderID.Id, SecondSubFolderID.Id, StaticThirdSubFolderID, Message.Id, _Email);                                            

                            // Creating the destnation folders.
                            string[] MakeDestinationFolderPath = { ImportFolderPath, _Email, FirstFolderName, SecondFolderName };
                            var DestinationFolderPath = Path.Combine(MakeDestinationFolderPath);

                            // Calls the folder cleaner to remove empty folders.
                            FolderCleanerClass.GetFolders(DestinationFolderPath);

                            // Variable used to store all the accepted extentions.
                            string[] AcceptedExtentions = ConfigParam.AllowedExtentions;

                            // Initilizing the download folder path variable.
                            string PathFullDownloadFolder = string.Empty;

                            // Export switch
                            var MoveToExport = false;

                            // Counter for below attachment downloadign foreach.
                            int Count = 0;

                            // FolderNameRnd creates a 10 digit folder name. CheckFolder returns the download path.
                            // This has to be called here. Don't put it within the for loop or it will start calling this
                            // each time and make folder for each file. Also calling this out side of the extentions FOR loop.
                            // causes an exception error at the "DownloadFileExistTest" test due file not been available.
                            PathFullDownloadFolder = Path.Combine(GraphHelper.CheckFolders("Download"), GraphHelper.FolderNameRnd(10));

                            if (Message.Attachments.Count > 0)
                            {
                                foreach (var Attachment in Message.Attachments.Where(x => AcceptedExtentions.Contains(Path.GetExtension(x.Name.ToLower())) && x.Size > 10240 || (x.Name.ToLower().EndsWith(".pdf") && x.Size < 10240)))
                                {
                                    Count++; // Count the for each execution once complete triggers the move.

                                    WriteLogClass.WriteToLog(3, "Collection check succeeded ...");

                                    // Should mark and download the itemattacment which is the correct attachment.
                                    var TrueAttachment = await graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                                                    .ChildFolders[$"{FirstSubFolderID.Id}"]
                                                    .ChildFolders[$"{SecondSubFolderID.Id}"]
                                                    .Messages[$"{Message.Id}"]
                                                    .Attachments[$"{Attachment.Id}"]
                                                    .Request()
                                                    .GetAsync();

                                    // Details of the attachment.
                                    var TrueAttachmentProps = (FileAttachment)TrueAttachment;
                                    string TrueAttachmentName = TrueAttachmentProps.Name.Replace(@"\", "").Replace("/", " ").Replace(":", "").Replace(";","");
                                    byte[] TruAttachmentBytes = TrueAttachmentProps.ContentBytes;

                                    // Extracts the extention of the attachment file.
                                    var AttExtention = Path.GetExtension(TrueAttachmentName).ToLower();

                                    // Check the name for "\", "/", and "c:".
                                    // If matched name is passed through the below function to normalize it.
                                    Regex MatchChar = new Regex(@"[\\\/c:]");

                                    if (MatchChar.IsMatch(TrueAttachmentName.ToLower()))
                                    {
                                        TrueAttachmentName = Path.GetFileName(TrueAttachmentName);
                                    }

                                    WriteLogClass.WriteToLog(3, $"Starting attachment download from {Message.Subject} ....");

                                    // Saves the file to the local hard disk.
                                    GraphHelper.DownloadAttachedFiles(PathFullDownloadFolder, TrueAttachmentName, TruAttachmentBytes);

                                    WriteLogClass.WriteToLog(3, $"Downloaded attachments from {Message.Subject}   ....");
                                    WriteLogClass.WriteToLog(3, $"Attachment name {TrueAttachmentName}");

                                    // Creating the metdata file.
                                    //var FileFlag = CreateMetaDataXml.GetToEmail4Xml(graphClient, FirstSubFolderID.Id, SecondSubFolderID.Id, StaticThirdSubFolderID, Message.Id, _Email, PathFullDownloadFolder, TrueAttachmentName);
                                }
                            }

                            if (Count > 0 && System.IO.Directory.Exists(PathFullDownloadFolder) && System.IO.Directory.EnumerateFiles(PathFullDownloadFolder, "*", SearchOption.AllDirectories).Any())
                            {
                                WriteLogClass.WriteToLog(3, "Moving downloaded files to local folder ....");

                                // Moves the downloaded files to destination folder. This would create the folder path if it's missing.
                                if (GraphHelper.MoveFolder(PathFullDownloadFolder, DestinationFolderPath))
                                {
                                    WriteLogClass.WriteToLog(3, "File/s moved successfully ....");
                                    MoveToExport = true;
                                }
                                else
                                {
                                    WriteLogClass.WriteToLog(3, "File/s was/were not moved successfully ....");
                                }
                            }

                            if (MoveToExport)
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
                                            if (await GraphHelper.MoveEmails(FirstSubFolderID.Id, SecondSubFolderID.Id, StaticThirdSubFolderID, MessageID, MoveDestinationID, _Email))
                                            {
                                                WriteLogClass.WriteToLog(3, $"Email {Message.Subject} moved to export folder ...");
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

                            if (Message.Attachments.Count() == 0 || Count == 0)
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
                                            var AttachmentStatus = 0; // No Attachment
                                            var StatiicSecondFolderID = string.Empty;

                                            if (Count == 0)
                                            {
                                                AttachmentStatus = 1; // Extension not accepted.
                                            }

                                            // Email is beeing forwarded.
                                            var ForwardDone = await GraphHelper.ForwardEmtpyEmail(FirstSubFolderID.Id, StatiicSecondFolderID, ErrorFolderId, MessageID2, _Email, AttachmentStatus);

                                            // After forwarding checks if the action returned true.
                                            // Item2 is the bool value returned.
                                            // Item1 is the maile address.
                                            if (ForwardDone.Item2)
                                            {
                                                WriteLogClass.WriteToLog(3, $"Email forwarded to {ForwardDone.Item1}  ....");
                                            }
                                            else
                                            {
                                                WriteLogClass.WriteToLog(3, $"Email not forwarded to {ForwardDone.Item1}  ....");
                                            }

                                            // Moves the empty emails to error folder once forwarding is done.
                                            if (await GraphHelper.MoveEmails(FirstSubFolderID.Id, SecondSubFolderID.Id, StaticThirdSubFolderID, MessageID2, ErrorFolderId, _Email) && ForwardDone.Item2)
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
            catch (Exception ex)
            {
                WriteLogClass.WriteToLog(1, $"Exception at end of main foreach 2level: {ex.Message}");
            }
        }
    }
}