using System.Text.RegularExpressions;
using WriteLog;

namespace FolderCleaner
{
    internal class FolderCleanerClass
    {
        public static void GetFolders(string FolderPath)
        {
            if (Directory.Exists(FolderPath))
            {                
                var LastFolderName = FolderPath.Split(Path.DirectorySeparatorChar).Last();

                Regex LastFolderNameMatch = new Regex(@"[0-9a-z]+@efakturamottak\.no");

                if (LastFolderNameMatch.IsMatch(LastFolderName.ToLower()))
                {
                    var CleaningFolderPath = Directory.GetParent(FolderPath);
                    FolderPath = CleaningFolderPath!.FullName;
                }              

                WriteLogClass.WriteToLog(3, $"Cleaning folder path {FolderPath} ....");

                string[] FolderList = Directory.GetDirectories(FolderPath, "*.*", SearchOption.AllDirectories);

                DeleteFolders(FolderList, LastFolderName);
            }
            else
            {
                WriteLogClass.WriteToLog(3, "Folder path does not exsits ....");
            }
        }

        private static void DeleteFolders(string[] _FolderList, string _LastFolder)
        {
            foreach (string _Folder in _FolderList)
            {
                if (Directory.Exists(_Folder))
                {                    
                    if (Directory.GetFiles(_Folder, "*.*", SearchOption.AllDirectories).Length == 0)
                    {
                        try
                        {                            
                            var RmvFolderName = _Folder.Split(Path.DirectorySeparatorChar).Last();
                            Regex LastFolderNameMatch = new Regex(@"[0-9a-z]+@efakturamottak\.no");

                            if (LastFolderNameMatch.IsMatch(_LastFolder.ToLower()))
                            {
                                Directory.Delete(_Folder,true);
                                WriteLogClass.WriteToLog(3, $"Folder {RmvFolderName} .... deleted");
                            }
                            else
                            {
                                var RmvdFolderName = _Folder.Split(Path.DirectorySeparatorChar).Last();
                                Directory.Delete(_Folder, false);
                                WriteLogClass.WriteToLog(3, $"Folder {RmvdFolderName} .... deleted");
                            }
                        }
                        catch (IOException ex)
                        {
                            WriteLogClass.WriteToLog(2, $"Exception at folder delete: {ex.Message}");
                        }

                    }
                    else
                    {
                        WriteLogClass.WriteToLog(3, $"Folder not empty .... not deleted");
                    }
                }
            }
        }
    }
}
