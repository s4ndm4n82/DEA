using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ReadSettings
{
    public class ReadSettingsClass
    {
        // Read the dea.conf file and adds the line into DeaConfig array.
        public string[] ReadConfig()
        {
            string[] DeaConfigs = { };

            try
            {
                var ConfigFolderPath = Directory.GetCurrentDirectory(); // Working Directory.
                var ConfigFileName = "dea.conf"; // Config file name.
                var ConfigFileFullPath = Path.Combine(ConfigFolderPath, ConfigFileName); // Makes the config file path.
                DeaConfigs = File.ReadAllLines(ConfigFileFullPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception at reading the conf file: {0}", ex.Message);
            }

            return DeaConfigs;
        }

        private string ReturnConfValue(string SearchTerm)
        {
            var ConfigValue = string.Empty;

            try
            {
                string[] ConfigData = ReadConfig();
                int pos = Array.FindIndex(ConfigData, row => row.Contains($"{SearchTerm}"));
                ConfigValue = ConfigData[pos].Replace(" ", "").Split('=').Last();                
            }
            catch (Exception ex)
            {
                Console.WriteLine("Exception at caonf value return: {0}", ex.Message);
            }

            return ConfigValue;
        }

        public bool DateFilter
        {
            get
            {
                bool dFilter;

                return dFilter = bool.Parse(ReturnConfValue("DateFilter"));                
            }
        }
        public int MaxLoadEmails
        {
            get
            {
                int EmailLoad;

                return EmailLoad = int.Parse(ReturnConfValue("MaxLoadEmails"));
            }
        }
        public string[]? UserAccounts
        {
            get
            {
                string[]? Emails;

                Emails = ReturnConfValue("UserAccounts").Split(",");
                List<string> EmailList = new List<string>(Emails.Length);
                EmailList.AddRange(Emails);
                return EmailList.ToArray();
            }
        }
        public string ImportFolderLetter
        {
            get
            {
                string? impFld;

                return impFld = ReturnConfValue("ImportFolderLetter");
            }
        }
        public string ImportFolderPath
        {
            get
            {
                string? impNme;

                return impNme = ReturnConfValue("ImportFolderPath");
            }
        }
        public string[] AllowedExtentions
        {
            get
            {
                string[]? exts;

                exts = ReturnConfValue("AllowedExtentions").Split(",");
                List<string> extList = new List<string>(exts.Length);
                extList.AddRange(exts);
                return exts.ToArray();
            }
        }
    }
}
