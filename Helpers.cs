using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Reflection;
using System.IO;
using System.Text.RegularExpressions;
using System.Net;
using OfficeOpenXml;

namespace OkCupidAutoBot
{
    class Helpers
    {
        public static string AssemblyDirectory
        {
            get
            {
                string codeBase = Assembly.GetExecutingAssembly().CodeBase;
                UriBuilder uri = new UriBuilder(codeBase);
                string path = Uri.UnescapeDataString(uri.Path);

                return Path.GetDirectoryName(path);
            }
        }

        public static int ReturnNumber(string subjectString, int group)
        {
            //string resultString = String.Empty;
            Regex regexObject = new Regex("(\\d+).*?(\\d+)", RegexOptions.IgnoreCase|RegexOptions.Singleline);
            Match m = default(Match);

            try
            {
                m = regexObject.Match(subjectString);
            }
            catch (ArgumentException exc)
            {
                Console.WriteLine(exc.Message);
            }

            return int.Parse(m.Groups[group].ToString());
        }

        public static string GetSafeFileName(string fileName)
        {
            return String.Join("_", fileName.Split(Path.GetInvalidFileNameChars()));
        }

        public static Cookie GetSpecificCookie(CookieContainer cookieJar, Uri domain, string cookieName)
        {
            CookieCollection okcCookies = cookieJar.GetCookies(domain);
            Cookie cookie = okcCookies[cookieName];

            return cookie;
        }

        public static string TransformToSqlParams(string fieldValues)
        {
            List<string> fields = fieldValues.Split(',').ToList();
            var fieldParams = fields.Select(q => "@" + q).ToList();
            string result = String.Join(",", fieldParams);

            return result;
        }

        public static bool HasLetters(string subjectString)
        {
            if (Regex.IsMatch(subjectString, @"
                # Match string having one letter and one digit (min).
                    \A                        # Anchor to start of string.
                      (?=[^0-9]*[0-9])        # at least one number and
                      (?=[^A-Za-z]*[A-Za-z])  # at least one letter.
                      \w+                     # Match string of alphanums.
                    \Z                        # Anchor to end of string.
                    ",
                    RegexOptions.IgnorePatternWhitespace))
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public static T ParseEnum<T>(string value)
        {
            return (T)Enum.Parse(typeof(T), value, true);
        }

        public static bool DoesSheetExist(FileInfo outputPath, string sheetName)
        {
            using (ExcelPackage package = new ExcelPackage(outputPath))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];

                if (worksheet == null)
                    return false;
                else
                    return true;
            }
        }
    }
}
