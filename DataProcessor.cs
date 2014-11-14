using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using HtmlAgilityPack;
using System.Web;
using System.Text.RegularExpressions;
using System.Globalization;
using System.Reflection;

namespace OkCupidAutoBot
{
    class DataProcessor
    {
        public ProfileDetails ProcessDetails(HtmlNode detailsNode, string herId, string userName, int age, string location, int matchPercent, int enemy)
        {
            //xpaths
            string detailsxPath = "./dl";
            HtmlNodeCollection nodeChildren = detailsNode.SelectNodes(detailsxPath);

            List<string> detailsList = CleanDetails(nodeChildren);

            DateTime lastOnline = ParseLastOnline(detailsList[0]);

            ProfileDetails details = new ProfileDetails();
            details.ProfileId = herId;
            details.Username = userName;
            details.Age = age;
            details.Location = location;
            details.Percentage = matchPercent;
            details.Enemy = enemy;
            details.LastOnline = lastOnline;
            details.Orientation = detailsList[1];
            details.Ethnicity = detailsList[2];
            details.Height = detailsList[3];
            details.BodyType = detailsList[4];
            details.Diet = detailsList[5];
            details.Smokes = detailsList[6];
            details.Drinks = detailsList[7];
            details.Drugs = detailsList[8];
            details.Religion = detailsList[9];
            details.Sign = detailsList[10];
            details.Education = detailsList[11];
            details.Job = detailsList[12];
            details.Income = detailsList[13];
            details.RelationshipStatus = detailsList[14];
            details.RelationshipType = detailsList[15];
            details.Offspring = detailsList[16];
            details.Pets = detailsList[17];
            details.Speaks = detailsList[18];
            details.CreatedBy = Program.Constants.AppName;
            details.CreatedDt = DateTime.Now;

            return details;
        }

        public List<string> CleanDetails(HtmlNodeCollection detailCollection)
        {
            List<string> resultDetail = new List<string>();

            //get the 'dd' items of each 'dl' list item
            var detailCollectionInner = detailCollection.Select(p => p.SelectSingleNode("./dd"));

            foreach (HtmlNode detail in detailCollectionInner)
            {
                string cleanedText = HttpUtility.HtmlDecode(detail.InnerText).Trim();
                resultDetail.Add(cleanedText);
            }

            return resultDetail;
        }

        public DateTime ParseLastOnline(string lastOnline)
        {
            string result = String.Empty;

            string[] lastOnlineArray = lastOnline.Split(' ');

            if (lastOnline == "Online now!")
            {
                return DateTime.Now;
            }
            else if (lastOnlineArray[0].ToLower() == "today" || lastOnlineArray[0].ToLower() == "yesterday")
            {
                DateTime rightNow = DateTime.Now;

                string today = rightNow.ToString("yyyy-MM-dd");

                string timeRegex = "^(?:[01]?[0-9]|2[0-3]):[0-5][0-9](?:[Aa]|[Pp])[Mm]";
                Match theTime = Regex.Match(lastOnlineArray[2], timeRegex);
                string timeString = theTime.Groups[0].Value;

                DateTime dateTimeLastOnline = DateTime.ParseExact(timeString, "h:mmtt", CultureInfo.InvariantCulture);
                string timeLastOnline = dateTimeLastOnline.ToString("hh:mm tt");

                if (lastOnlineArray[0].ToLower() == "today")
                {
                    result = today + " " + timeLastOnline;
                }
                else
                {
                    DateTime yesterdayDateTime = rightNow.AddDays(-1.0);
                    string yesterday = yesterdayDateTime.ToString("yyyy-MM-dd");

                    result = yesterday + " " + timeLastOnline;
                }
            }
            else
            {
                string monthRegex = "((?:Jan(?:uary)?|Feb(?:ruary)?|Mar(?:ch)?|Apr(?:il)?|May|Jun(?:e)?|Jul(?:y)?|Aug(?:ust)?|Sep(?:tember)?|Sept|Oct(?:ober)?|Nov(?:ember)?|Dec(?:ember)?))";
                string filler = ".*?"; //non-greedy match filler
                string dayRegex = "((?:(?:[0-2]?\\d{1})|(?:[3][01]{1})))(?![\\d])";

                Regex r = new Regex(monthRegex + filler + dayRegex, RegexOptions.IgnoreCase | RegexOptions.Singleline);
                Match m = r.Match(lastOnline);

                if (m.Success)
                {
                    result = m.Groups[1].ToString() + " " + m.Groups[2].ToString() + " " + DateTime.Now.Year;
                }
                else
                    throw new Exception("No match found!");
                
            }
            
            return DateTime.Parse(result);
        }

        public Essays ProcessEssays(HtmlNode essaysNode, string herId)
        {
            Essays essays = new Essays();
            essays.ProfileId = herId;
            essays.CreatedBy = Program.Constants.AppName;
            essays.CreatedDt = DateTime.Now;
            essays.what_i_want = HttpUtility.HtmlDecode(essaysNode.SelectSingleNode("//div[contains(@class,'what_i_want')]").InnerText).Trim();

            Type type = typeof(Essays);
            foreach (PropertyInfo pInfo in type.GetProperties().Where(p => p.Name.Contains("essay")).ToList())
            {
                //if(pInfo.Name.Contains("essay"))
                //{
                int essayIndex = int.Parse(pInfo.Name.Substring(pInfo.Name.Length - 1, 1));
                string essayXpath = String.Format("//div[@id='essay_text_{0}']", essayIndex);
                string essayString = default(string);
                string cleanString = default(string);

                HtmlNode essayNode = essaysNode.SelectSingleNode(essayXpath);

                if (essayNode != null)
                {
                    essayString = essayNode.InnerText;
                    cleanString = HttpUtility.HtmlDecode(essayString).Trim();
                }

                pInfo.SetValue(essays, cleanString, null);
                //}
            }

            return essays;
        }
    }
}
