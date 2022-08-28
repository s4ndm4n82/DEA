using Microsoft.Graph;
using System.Xml;
using System.Text;
using WriteLog;

namespace CreateMetadataFile
{
    internal class CreateMetaDataXml
    {
        public static bool WriteMetadataXml(string ToEmail, string SavePath)
        {
            // TODO 1 : Create a funstion to get the to email address from emails and pass it to here.
            var XmlSaveFile = Path.Combine(SavePath, "Metadata.xml");
            var XmlSaveSwitch = false;

            XmlWriterSettings WriterSettings = new XmlWriterSettings();
            WriterSettings.Indent = true;
            WriterSettings.IndentChars = (" ");
            WriterSettings.CloseOutput = true;
            WriterSettings.OmitXmlDeclaration = true;
            WriterSettings.Encoding = Encoding.UTF8;
            
            if (!string.IsNullOrEmpty(ToEmail) && !string.IsNullOrEmpty(SavePath))
            {
                try
                {
                    using (XmlWriter FileWriter = XmlWriter.Create(XmlSaveFile, WriterSettings))
                    {
                        FileWriter.WriteStartDocument();
                        FileWriter.WriteStartElement("BaseTypeContainer");
                        FileWriter.WriteStartElement("BaseTypeObject");
                        FileWriter.WriteStartElement("Metadata");
                        FileWriter.WriteStartElement("Fields");
                        FileWriter.WriteStartElement("Field");
                        FileWriter.WriteAttributeString("Type", "Text");
                        FileWriter.WriteAttributeString("Label", "Email");
                        FileWriter.WriteElementString("Value", ToEmail);
                        FileWriter.WriteEndElement();
                        FileWriter.WriteEndDocument();
                        FileWriter.Flush();
                        FileWriter.Close();
                    }

                    if (System.IO.File.Exists(XmlSaveFile))
                    {
                        WriteLogClass.WriteToLog(3, $"Metdata file created at {XmlSaveFile} ....");
                        XmlSaveSwitch = true;
                    }
                    else
                    {
                        WriteLogClass.WriteToLog(1, $"Unable to create metadata file at {XmlSaveFile} ....");
                    }
                }
                catch (Exception ex)
                {
                    WriteLogClass.WriteToLog(1, $"Exception at Xml metadata file creation: {ex.Message}");
                }
            }

            return XmlSaveSwitch;
        }

        public static bool GetToEmail4Xml(GraphServiceClient graphClient, string SubFolderId1, string SubFolderId2, string SubFolderId3, string MessageID, string _Email, string _FolderPath)
        {
            var FileFlag = false;
            IEnumerable<Recipient> ToEmails;
            Task<Message> GetToEmail;
            
            try
            {
                if (string.IsNullOrEmpty(SubFolderId3))
                {
                    GetToEmail = graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                            .ChildFolders[$"{SubFolderId1}"]
                            .ChildFolders[$"{SubFolderId2}"]
                            .Messages[$"{MessageID}"]
                            .Request()
                            .Select(eml => new
                            {
                                eml.ToRecipients
                            })
                            .GetAsync();
                }
                else
                {
                    GetToEmail = graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                            .ChildFolders[$"{SubFolderId1}"]
                            .ChildFolders[$"{SubFolderId2}"]
                            .ChildFolders[$"{SubFolderId3}"]
                            .Messages[$"{MessageID}"]
                            .Request()
                            .Select(eml => new
                            {
                                eml.ToRecipients
                            })
                            .GetAsync();
                }

                ToEmails = GetToEmail.Result.ToRecipients.Where(x => x.EmailAddress.Address.Contains("@efakturamottak.no"));

                foreach (var ToEmail in ToEmails)
                {
                    if (!string.IsNullOrEmpty(ToEmail.EmailAddress.Address))
                    {
                        WriteLogClass.WriteToLog(3, $"Recipient email {ToEmail.EmailAddress.Address} extracted ...");
                        FileFlag = WriteMetadataXml(ToEmail.EmailAddress.Address, _FolderPath);
                    }                    
                }
            }
            catch (Exception ex)
            {
                WriteLogClass.WriteToLog(1, $"Exception at GetToEmail4Xml {ex.Message}");                
            }

            return FileFlag;
        }
    }
}
