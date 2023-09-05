namespace Truncate.Extensions
{
    internal static class TruncateTool
    {
        /// <summary>
        /// Returns the filename after trunating them to a shorter set of charaters.
        /// </summary>
        /// <param name="value"></param>
        /// <param name="maxLength"></param>
        /// <returns></returns>
        public static string TruncateString(this string value, int maxLength)
        {
            return value.Length > maxLength ? value.Substring(0, maxLength) : value;
        }
    }
}
