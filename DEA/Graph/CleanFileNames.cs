using System.Text.RegularExpressions;
using Truncate.Extensions;

namespace CleanFileNames
{
    internal class CleanFileNamesClass
    {
        public static string CleanFileName(string fileName, string fileExtension)
        {
            string cleanedFileName = fileName.Length > 40 ? FileNameCleaner(fileName.TruncateString(10)) : FileNameCleaner(fileName);

            // Making the full file name after cleaning it.
            string returnFileName = Path.ChangeExtension(cleanedFileName, fileExtension);

            return returnFileName;
        }

        private static string FileNameCleaner(string passedFileName)
        {
            // Strips the filename of invalid charaters and replace them with "_".
            string regexPattern = @"[\\~#%&*{}/:<>?|""-\.]";
            string replaceChar = "_";
            Regex regexCleaner = new(regexPattern, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
            string fileNameClean = Regex.Replace(regexCleaner.Replace(passedFileName, replaceChar), @"[\s]+", "");

            return fileNameClean;
        }
    }
}
