using System.IO;
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
                var CleaningFolderName = Directory.GetParent(Path.GetDirectoryName(FolderPath)!);             

                WriteLogClass.WriteToLog(3, $"Cleaning folder path {CleaningFolderName!.FullName} ....");

                string[] FolderList = Directory.GetDirectories(FolderPath, "*.*", SearchOption.AllDirectories);

                DeleteFolders(FolderList);
            }
            else
            {
                try
                {
                    WriteLogClass.WriteToLog(3, "Folder path does not exsits. Creating folder path ....");

                    Directory.CreateDirectory(FolderPath);

                    WriteLogClass.WriteToLog(3, $"Folder path created {FolderPath} ....");
                }
                catch (Exception ex)
                {
                    WriteLogClass.WriteToLog(3, $"Path: {FolderPath}");
                    WriteLogClass.WriteToLog(2, $"Exception at folder creation in folder cleaner class: {ex.Message}");
                }
                
            }
        }

        private static void DeleteFolders(string[] _FolderList)
        {
            foreach (string _Folder in _FolderList)
            {
                if (Directory.Exists(_Folder))
                {
                    if (Directory.GetFiles(_Folder, "*.*", SearchOption.AllDirectories).Length == 0)
                    {
                        try
                        {   
                            var RmvFolderPath = Directory.GetParent(_Folder);
                            var RmvFolderName = RmvFolderPath!.FullName.Split(Path.DirectorySeparatorChar).Last();
                            Regex LastFolderNameMatch = new Regex(@"Download");

                            if (!LastFolderNameMatch.IsMatch(RmvFolderName))
                            {
                                Directory.Delete(RmvFolderPath.FullName,true);
                                WriteLogClass.WriteToLog(3, $"Folder {RmvFolderName} .... deleted");
                            }
                            else
                            {
                                var RmvdFromDownload = _Folder.Split(Path.DirectorySeparatorChar).Last();
                                Directory.Delete(_Folder, false);
                                WriteLogClass.WriteToLog(3, $"Cleaning {RmvdFromDownload} ....");
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
