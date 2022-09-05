using Microsoft.Graph;
using System.Text.RegularExpressions;
using WriteLog;

namespace GetRecipientEmail
{
    internal class GetRecipientEmailClass
    {
        public static string GetRecipientEmail(GraphServiceClient graphClient, string SubFolderId1, string SubFolderId2, string SubFolderId3, string MessageID, string _Email)
        {
            var rEmail = string.Empty;
            IEnumerable<InternetMessageHeader> ToEmails;
            Task<Message> GetToEmail;

            try
            {
                if (string.IsNullOrEmpty(SubFolderId3) && string.IsNullOrEmpty(SubFolderId2))
                {
                    GetToEmail = graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                            .ChildFolders[$"{SubFolderId1}"]
                            .Messages[$"{MessageID}"]
                            .Request()
                            .Select(eml => new
                            {
                                eml.InternetMessageHeaders
                            })
                            .GetAsync();
                }
                else if (string.IsNullOrEmpty(SubFolderId3))
                {
                    GetToEmail = graphClient.Users[$"{_Email}"].MailFolders["Inbox"]
                            .ChildFolders[$"{SubFolderId1}"]
                            .ChildFolders[$"{SubFolderId2}"]
                            .Messages[$"{MessageID}"]
                            .Request()
                            .Select(eml => new
                            {
                                eml.InternetMessageHeaders
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
                                eml.InternetMessageHeaders
                            })
                            .GetAsync();
                }

                ToEmails = GetToEmail.Result.InternetMessageHeaders.Where(adrs => adrs.Value.Contains("@efakturamottak.no"));

                foreach (var ToEmail in ToEmails)
                {
                    if (!string.IsNullOrEmpty(ToEmail.Value))
                    {
                        string RegExString = @"[0-9a-z]+@efakturamottak\.no";
                        Regex RecivedEmail = new Regex(RegExString, RegexOptions.IgnoreCase);
                        var ExtractedEmail = RecivedEmail.Match(ToEmail.Value);

                        if (ExtractedEmail.Success)
                        {
                            rEmail = ExtractedEmail.Value.ToLower().Replace(" ","");
                            WriteLogClass.WriteToLog(3, $"Recipient email {rEmail} extracted ...");
                            break;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                WriteLogClass.WriteToLog(1, $"Exception at getting recipient email: {ex.Message}");
            }

            return rEmail;
        }
    }
}
