using System.Text.RegularExpressions;

namespace macros.Utils
{
    public static class StringCleaner
    {
        public static string Clean(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            string result = value.Trim();

            // заменяем множественные пробелы на один
            result = Regex.Replace(result, @"\s+", " ");

            return result;
        }
    }
}
