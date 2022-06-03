using Microsoft.Graph;
using System.Diagnostics.CodeAnalysis;

namespace DEA2Levels
{
    internal class GraphHelper2Levels
    {
        public static async Task GetEmailsAttacmentsAccount([NotNull] GraphServiceClient graphClient, string _Email)
        {
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
                }
            }
        }
    }
}
