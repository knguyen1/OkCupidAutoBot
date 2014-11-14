using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Net;

namespace OkCupidAutoBot
{
    public class CookieAwareWebClient : WebClient
    {
        public CookieAwareWebClient() : this(new CookieContainer())
        {
        }

        public CookieAwareWebClient(CookieContainer cookieJar)
        {
            this.CookieContainer = cookieJar;
        }

        public CookieContainer CookieContainer { get; set; }

        protected override WebRequest GetWebRequest(Uri address)
        {
            WebRequest request = base.GetWebRequest(address);

            var castRequest = request as HttpWebRequest;
            if (castRequest != null)
            {
                //form the headers, make yourself look like chrome
                castRequest.CookieContainer = this.CookieContainer;
                castRequest.Accept = "text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,*/*;q=0.8";
                castRequest.UserAgent = "Mozilla/5.0 (Windows NT 6.1; WOW64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/34.0.1847.131 Safari/537.36";
                castRequest.Headers.Add("Accept-Encoding", "gzip,deflate,sdch");
                castRequest.Headers.Add("Accept-Language", "en-US,en;q=0.8");
                castRequest.Referer = "http://www.okcupid.com/match?timekey=1&matchOrderBy=SPECIAL_BLEND&use_prefs=1&discard_prefs=1&count=18";

                //automatic decompression
                castRequest.AutomaticDecompression = DecompressionMethods.GZip | DecompressionMethods.Deflate;
            }

            return request;
        }
    }
}
