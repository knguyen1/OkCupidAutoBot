using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;
using System.Collections.Specialized;
using HtmlAgilityPack;
using System.Collections;
using System.Text.RegularExpressions;
using System.Data.SqlClient;
using System.Threading;
using System.Reflection;
using OfficeOpenXml;
using Newtonsoft.Json.Linq;

namespace OkCupidAutoBot
{
    public class Program
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

        public class Constants
        {
            public static string matchResultDivId { get { return "match_results"; } }
            public static string matchCardClass { get { return "match_card"; } }
            public static string matchCardText { get { return "match_card_text"; } }
            public static string profileDetails { get { return "profile_details"; } }
            public static string essaysColumn { get { return "main_column"; } }
            public static string AppName { get { return System.Reflection.Assembly.GetExecutingAssembly().FullName.ToString(); } }
        }

        static void Main(string[] args)
        {
            //helpers
            DataProcessor ProcessData = new DataProcessor();
            SqlDataRepository SqlRepo = new SqlDataRepository();
            ExcelDataRepository ExcelRepo = new ExcelDataRepository();

            // pseudo-random number generator, will be used to decide how long to wait between each iterations
            // and whether or not to process the girl
            Random r = new Random();

            //write people to tables or excel sheet?
            int writeMethod = OkSettings.Default.sqlOrExcel;

            //excel file variables
            string outputDir = AssemblyDirectory + @"\excel\";
            string girlsOutput = outputDir + OkSettings.Default.excelFile;
            FileInfo profileExcel = new FileInfo(girlsOutput);

            // get the base address
            CookieAwareWebClient client = new CookieAwareWebClient();
            client.BaseAddress = OkSettings.Default.baseUri;

            CookieContainer cookieJar = client.CookieContainer;

            // establish login data
            var loginData = new NameValueCollection();
            loginData.Add("username", OkSettings.Default.username.ToLower());
            loginData.Add("password", OkSettings.Default.password);
            loginData.Add("okc_api", "1");

            //begin login, will return a byte[] array that we'll have to parse to string
            byte[] response = client.UploadValues("/login", "POST", loginData);
            string loginResponse = System.Text.Encoding.Default.GetString(response);

            //parse the json string and get login status
            JObject loginJson = JObject.Parse(loginResponse);
            int loginStatus = int.Parse(loginJson["status"].ToString());
            
            //login status
            //0: successful
            //104: bad username and password
            //107: lol, your ass is banned
            if (loginStatus == 0)
            {
                Console.WriteLine("{0} :: {1} :: Login successful.", DateTime.Now.ToShortTimeString(), OkSettings.Default.username.ToLower());
            }
            else
            {
                Console.WriteLine("{0} :: {1} :: Login failed.  Press any key to continue.", DateTime.Now.ToShortTimeString(), OkSettings.Default.username.ToLower());

                if (loginStatus == 107)
                    Console.WriteLine("Account banned!");

                Console.ReadLine();
                return;
            }

            //get the session id from the json string
            string sessionId = loginJson["userid"].ToString();

            ////NOM NOM NOM cookies... this is the log-winded way to get sessionId, from the cookie
            //Uri okcUri = new Uri(client.BaseAddress);
            //Cookie sessionCookie = Helpers.GetSpecificCookie(cookieJar, okcUri, "session");

            ////get the sessionId from the cookie
            //Match regexResult = Regex.Match(sessionCookie.Value.ToString(), @"^(\d+)%(\d+)");
            //string sessionId = regexResult.Groups[1].Value;

            //bool keepProcessing = default(bool);
            int girlsProcessed = 0;

            List<ProfileDetails> listOfGirls = new List<ProfileDetails>();
            List<Essays> listOfEssays = new List<Essays>();

            do
            {
                //now you are logged in and can get pages
                //search settings take the saved settings of whatever it was set on the UI
                //you can also modify the matchesQueryString variable to fit your needs
                Console.WriteLine("Downloading a new set of girls...");
                string matches = client.DownloadString(OkSettings.Default.matchesQueryString);

                List<string> processedGirls = new List<string>();

                //read from SQL database and add girls names to the already processed list
                //we will use this later to decide whether or not to visit the girl again
                if (writeMethod == 1)
                {
                    Console.WriteLine("Using write method 1, will write to SQL server.");

                    using (SqlConnection conn = new SqlConnection(OkSettings.Default.sqlConnection))
                    {
                        string query = String.Format("SELECT DISTINCT Username FROM {0}", OkSettings.Default.detailsTable);

                        SqlCommand command = conn.CreateCommand();
                        command.CommandType = System.Data.CommandType.Text;
                        command.CommandText = query;

                        if (conn.State != System.Data.ConnectionState.Open)
                            conn.Open();

                        using (SqlDataReader reader = command.ExecuteReader())
                        {
                            while (reader.Read())
                            {
                                processedGirls.Add(reader.GetString(0));
                            }
                        }
                    }
                }
                else if (writeMethod == 2)
                {
                    Console.WriteLine("Using write method 1, will write to excel sheet.");

                    using (ExcelPackage package = new ExcelPackage(profileExcel))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets["GirlsDetails"];

                        if (worksheet != null)
                        {
                            int lastRow = worksheet.Dimension.End.Row;
                            for (int i = 1; i <= lastRow; i++)
                            {
                                //skip the header row
                                int currentRow = i + 1;

                                string cellValue = (string)worksheet.Cells[currentRow, 3].Value;

                                //check if the cell has a value, if not break the loop
                                if (cellValue == null)
                                    break;

                                processedGirls.Add(cellValue);
                            }

                            //get unique girls only
                            processedGirls = processedGirls.Distinct(StringComparer.CurrentCultureIgnoreCase).ToList();
                        }
                    }
                }

                // convert to HTML document
                HtmlDocument matchesHtmlDoc = new HtmlDocument();
                matchesHtmlDoc.LoadHtml(matches);

                //to get HTML nodes, you need to define x paths
                //more info here: http://www.w3schools.com/xpath/xpath_nodes.asp
                string matchesXpath = String.Format("//div[@id='{0}']/div", Constants.matchResultDivId);

                // THIS IS THE MEAT OF EVERYTHING
                // this section gets all the match cards and places them into a collection of nodes called 'people'
                // we will then iterate over the nodes and parse the girls, one by one
                HtmlNodeCollection people = matchesHtmlDoc
                    .DocumentNode
                    .SelectNodes(matchesXpath);

                //throttles calls using RateGate
                //RateGate takes two arguments: number of occurrence and TimeSpan
                //ex: RateGate(1, TimeSpan.FromSeconds(3)) will allow one call per three seconds
                using (var RateGate = new RateGate(1, TimeSpan.FromSeconds(5)))
                {
                    foreach (HtmlNode person in people)
                    {
                        string matchCardXpath = String.Format("./div[contains(@class,'{0}')]", Constants.matchCardClass);
                        string matchCardTextXpath = String.Format("{0}/div[@class='{1}']", matchCardXpath, Constants.matchCardText);
                        string profileInfoXpath = String.Format("{0}/div[@class='profile_info']", matchCardTextXpath);
                        string userInfoXpath = String.Format("{0}/div[@class='userinfo']", profileInfoXpath);

                        // find person info, including name and link to picture
                        string personLinkXpath = String.Format("{0}/a", matchCardXpath);
                        HtmlNode personLinkNode = person.SelectSingleNode(personLinkXpath);

                        string userName = personLinkNode.Attributes["data-username"].Value;

                        //check if username is already in excel/DB or already processed and stored in the running list
                        if (processedGirls.Contains(userName) || listOfGirls.Where(p => p.Username.Contains(userName)).FirstOrDefault() != null)
                        {
                            if (r.Next(0, 2) == 0)
                            {
                                Console.WriteLine("{0} :: {1} :: Processed her... Skipping.", DateTime.Now.ToString("h:mm:ss tt"), userName);
                                continue;  //don't process her again
                            }

                            //if the inner IF statement is not triggered, we'll visit the girl again
                            //we just won't parse her data...
                            Console.WriteLine("{0} :: {1} :: Processed her... Visiting her again.", DateTime.Now.ToString("h:mm:ss tt"), userName);
                        }

                        girlsProcessed++;

                        // userinfo xpaths
                        string profileAgeXpath = String.Format("{0}/span[@class='age']", userInfoXpath);
                        string profileLocationXpath = String.Format("{0}/span[@class='location']", userInfoXpath); ;
                        string matchPercentageXpath = String.Format("{0}/div[contains(@class,'percentages')]", matchCardTextXpath); ;

                        // userinfo nodes
                        HtmlNode profileAgeNode = person.SelectSingleNode(profileAgeXpath);
                        HtmlNode profileLocationNode = person.SelectSingleNode(profileLocationXpath);
                        HtmlNode matchPercentageNode = person.SelectSingleNode(matchPercentageXpath);

                        // userinfo text
                        int age = int.Parse(profileAgeNode.InnerText);
                        string location = profileLocationNode.InnerText;
                        string matchPercentageFullText = matchPercentageNode.InnerText;
                        int matchPercent = Helpers.ReturnNumber(matchPercentageFullText, 1);
                        int enemy = Helpers.ReturnNumber(matchPercentageFullText, 2);

                        string personProfileLink = personLinkNode.Attributes["href"].Value;
                        string personProfilePic = personLinkNode.Attributes["data-image-url"].Value;

                        Console.WriteLine("{0} :: user {1} :: {2}/{3}", DateTime.Now.ToString("h:mm:ss tt"), userName, people.IndexOf(person).ToString(), people.Count().ToString());

                        string thisProfile = default(string);

                        ////OPTIONAL: You can define a match threshold if you'd like
                        ////un-comment these lines if you want to skip girls that do not satisfy the threshhold
                        //if (matchPercent < OkSettings.Default.matchThreshhold)
                        //    continue;

                        //throttles the calls using the rate gate
                        if (OkSettings.Default.callThrottling)
                            RateGate.WaitToProceed();

                        try
                        {
                            thisProfile = client.DownloadString(personProfileLink);
                        }
                        catch (WebException exc)
                        {
                            Console.WriteLine(exc.Message);
                        }

                        if (thisProfile != null)
                        {
                            //convert to html document
                            HtmlDocument profileHtml = new HtmlDocument();
                            profileHtml.LoadHtml(thisProfile);
                            HtmlNode profileDocNode = profileHtml.DocumentNode;

                            //get her user id
                            string ratingXpath = "//ul[@id='personality-rating']/li/a[contains(@class, 'star')]";
                            HtmlNode ratingNode = profileDocNode.SelectSingleNode(ratingXpath);

                            string ratingLink = ratingNode.Attributes["href"].Value;
                            string linkRegex = @"[0-9], '(\d+)',";

                            Match idRegex = Regex.Match(ratingLink, linkRegex);
                            string herId = idRegex.Groups[1].Value;

                            //details and essays xpaths
                            string detailsXpath = String.Format("//div[@id='{0}']", Constants.profileDetails);
                            string essaysXpath = String.Format("//div[@id='{0}']", Constants.essaysColumn);

                            //details and essays nodes
                            HtmlNode detailsNode = profileDocNode.SelectSingleNode(detailsXpath);
                            HtmlNode essaysNode = profileDocNode.SelectSingleNode(essaysXpath);

                            //process the data
                            ProfileDetails profileDetails = ProcessData.ProcessDetails(detailsNode, herId, userName, age, location, matchPercent, enemy);
                            Essays essay = ProcessData.ProcessEssays(essaysNode, herId);

                            //if the user has NOT been processed, AND the running list does NOT contain the user name
                            //add her
                            if (!processedGirls.Contains(userName) && listOfGirls.Where(p => p.Username.Contains(userName)).FirstOrDefault() == null)
                            {
                                Console.WriteLine("Adding...");
                                listOfGirls.Add(profileDetails);
                                listOfEssays.Add(essay);
                            }

                            //get the current rating
                            string currentRatingXpath = "//li[contains(@id,'current-personality')]";
                            HtmlNode _currentRating = profileHtml.DocumentNode.SelectSingleNode(currentRatingXpath);
                            HtmlAttribute styleAttr = _currentRating.Attributes["style"];

                            //has she been rated?
                            bool hasVoted = default(bool);

                            //make sure it has the style attr
                            if (styleAttr != null)
                            {
                                string currentRating = Regex.Match(styleAttr.Value, @"\d+").Value;

                                //if current rating is > 0, then you already voted
                                if (Int32.Parse(currentRating) > 0)
                                    hasVoted = true;
                                else
                                    hasVoted = false;
                            }

                            //only initiate if girl has not been voted
                            if (!hasVoted)
                            {
                                //randomly rate her or not at a rate of 1/2 girls per encounter
                                if (r.Next(101) < 50)
                                {
                                    //randomly generate her score between 3-5... we don't wanna be always rating her a 5.. like a robot ;)
                                    int herScore = r.Next(3, 6);

                                    NameValueCollection voteForm = new NameValueCollection();
                                    voteForm.Add("voterid", sessionId);
                                    voteForm.Add("target_userid", herId);
                                    voteForm.Add("type", "vote");
                                    voteForm.Add("target_objectid", "0");
                                    voteForm.Add("vote_type", "personality");
                                    voteForm.Add("score", herScore.ToString());

                                    try
                                    {
                                        client.UploadValues("https://www.okcupid.com/vote_handler", "POST", voteForm);
                                    }
                                    catch (WebException exc)
                                    {
                                        Console.WriteLine(exc.Message);
                                    }

                                    Console.WriteLine("Voted her: {0}", herScore);
                                }
                                else
                                    Console.WriteLine("Didn't vote for her...");
                            }
                            else
                                Console.WriteLine("Already voted for her...");
                        }

                        ////you can also use Thread.Sleetp to throttle calls (not recommended)
                        ////randomly wait 5-10 seconds between each call
                        //if (OkSettings.Default.callThrottling)
                        //{
                        //    int seconds = r.Next(5, 11);
                        //    TimeSpan waitTime = TimeSpan.FromSeconds(seconds);

                        //    Console.WriteLine("Waiting... {0} seconds", seconds);
                        //    Thread.Sleep(waitTime);
                        //}

                    } //foreach girls loop

                } //RateGate

                Console.WriteLine("Girls processed: {0}", girlsProcessed);
            }
            while (girlsProcessed < OkSettings.Default.girlsPerSession);

            //save the data
            // 1: SQL database write method
            // 2: Excel worksheet writemethod
            if (writeMethod == 1)
            {
                SqlRepo.SaveItems<ProfileDetails>(listOfGirls, OkSettings.Default.detailsTable);
                SqlRepo.SaveItems<Essays>(listOfEssays, OkSettings.Default.essaysTable);
            }
            else if (writeMethod == 2)
            {
                ExcelRepo.SaveDetails(listOfGirls, profileExcel);
                ExcelRepo.SaveEssays(listOfEssays, profileExcel);
            }
        }
    }
}
