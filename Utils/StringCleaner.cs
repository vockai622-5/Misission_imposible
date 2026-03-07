using System.Linq;
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

            result = Regex.Replace(result, @"[\u0000-\u001F\u007F\u00A0\u200B\u200C\u200D\uFEFF]", "");
            result = Regex.Replace(result, @"\s+", " ");

            return result.Trim();
        }

        public static string CleanName(string value)
        {
            var cleaned = Clean(value);
            if (string.IsNullOrEmpty(cleaned))
                return string.Empty;

            cleaned = Regex.Replace(cleaned, @"[^a-zA-Zа-яА-ЯёЁ\s\-]", "");
            cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim();
            cleaned = Regex.Replace(cleaned, @"\-{2,}", "-");

            if (string.IsNullOrEmpty(cleaned))
                return string.Empty;

            return CapitalizeParts(cleaned);
        }

        public static string CleanLogin(string value)
        {
            var cleaned = Clean(value);
            if (string.IsNullOrEmpty(cleaned))
                return string.Empty;

            cleaned = cleaned.Replace(" ", "");
            cleaned = cleaned.ToLowerInvariant();
            cleaned = Regex.Replace(cleaned, @"[^a-z0-9._\-]", "");

            return cleaned;
        }

        public static string CleanGroupName(string value)
        {
            var cleaned = Clean(value);
            if (string.IsNullOrEmpty(cleaned))
                return string.Empty;

            cleaned = Regex.Replace(cleaned, @"[^a-zA-Zа-яА-ЯёЁ0-9_\-\s]", "");
            cleaned = Regex.Replace(cleaned, @"\s+", " ").Trim();

            return cleaned;
        }

        public static string BuildShortName(string lastName, string firstName, string middleName)
        {
            lastName = CleanName(lastName);
            firstName = CleanName(firstName);
            middleName = CleanName(middleName);

            if (string.IsNullOrWhiteSpace(lastName))
                return string.Empty;

            var firstInitial = GetInitial(firstName);
            var middleInitial = GetInitial(middleName);

            var parts = new[]
            {
                lastName,
                firstInitial,
                middleInitial
            }
            .Where(x => !string.IsNullOrWhiteSpace(x));

            return string.Join(" ", parts);
        }

        private static string GetInitial(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return string.Empty;

            return char.ToUpper(value[0]) + ".";
        }

        private static string CapitalizeParts(string value)
        {
            var words = value.Split(' ');
            for (int i = 0; i < words.Length; i++)
            {
                words[i] = CapitalizeHyphenated(words[i]);
            }

            return string.Join(" ", words);
        }

        private static string CapitalizeHyphenated(string word)
        {
            if (string.IsNullOrEmpty(word))
                return word;

            var parts = word.Split('-');
            for (int i = 0; i < parts.Length; i++)
            {
                if (parts[i].Length > 0)
                    parts[i] = char.ToUpper(parts[i][0]) + parts[i].Substring(1).ToLower();
            }

            return string.Join("-", parts);
        }
    }
}