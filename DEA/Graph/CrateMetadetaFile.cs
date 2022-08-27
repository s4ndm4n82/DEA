using Microsoft.Graph;
using WriteLog;
using System.Xml;
using System.Text;
using System.Linq;
using System.Text.RegularExpressions;

namespace CreateMetadataFile
{
    internal class CreateMetaDataXml
    {
        public static void WriteMetadataXml(string ToEmail)
        {
            // TODO 1 : Create a funstion to get the to email address from emails and pass it to here.

            XmlWriterSettings WriterSettings = new XmlWriterSettings();
            WriterSettings.Indent = true;
            WriterSettings.IndentChars = (" ");
            WriterSettings.CloseOutput = true;
            WriterSettings.OmitXmlDeclaration = true;
            WriterSettings.Encoding = Encoding.UTF8;

            using(XmlWriter FileWriter = XmlWriter.Create("Metadata.xml", WriterSettings))
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
        }

        public static async void GetToEmail4Xml(GraphServiceClient graphClient, string SubFolderId1, string SubFolderId2, string MessageID, string _Email)
        {
            try
            {
                var GetToEmail = await graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                            .ChildFolders[$"{SubFolderId1}"]
                            .ChildFolders[$"{SubFolderId2}"]
                            .Messages[$"{MessageID}"]
                            .Request()
                            .Select(eml => new
                            {
                                eml.ToRecipients
                            })
                            .GetAsync();

                foreach (var recipient in GetToEmail.ToRecipients.Select(x => x.EmailAddress))
                {
                    Regex SelectEmails = new Regex(@"^.+@efakturamottak.no$");
                    if (SelectEmails.IsMatch(recipient.Address))
                    {
                        WriteLogClass.WriteToLog(3, $"Email recipiant from XML -> {recipient.Address}");
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLogClass.WriteToLog(1, $"Exception at xml email get: {ex.Message}");
            }
        }
    }
}
