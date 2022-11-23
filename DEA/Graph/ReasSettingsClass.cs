using Newtonsoft.Json;
using WriteLog;

namespace ReadSettings
{
    internal class ReasSettingsClass
    {

        public class ReasSettingsObject
        {
            public Appsetting[]? AppSettings { get; set; }
            public Folderstoexclude[]? FoldersToExclude { get; set; }
        }

        public class Appsetting
        {
            public bool DateFilter { get; set; }
            public int MaxLoadEmails { get; set; }
            public string[]? UserAccounts { get; set; }
            public string? ImportFolderLetter { get; set; }
            public string? ImportFolderPath { get; set; }
            public string[]? AllowedExtentions { get; set; }
        }

        public class Folderstoexclude
        {
            public string[]? MainEmailFolders { get; set; }
            public object[]? SubEmailFolders { get; set; }
        }

        public static T ReadAppConfig<T>()
        {
            T? jsonFileData = default;

            try
            {
                using StreamReader fileData = new StreamReader(@".\appconfig.json");
                string jsonDataString = fileData.ReadToEnd();
                jsonFileData = JsonConvert.DeserializeObject<T>(jsonDataString);

                return jsonFileData!;
            }
            catch (Exception ex)
            {
                WriteLogClass.WriteToLog(3, $"Exception at Json settings reader: {ex.Message}");
                return jsonFileData!;
            }
        }

    }
}
