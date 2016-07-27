using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;

namespace WriteWord
{
    class HttpHelper
    {
        private static CookieContainer msgCookies = new CookieContainer();
        public static bool getResponseText(string url, out string responseText, string method = "GET", string data = "", string user_name = "", string password = "")
        {
            Uri uri = new Uri(url);
            HttpWebRequest request = (HttpWebRequest)WebRequest.Create(uri);
            if(user_name != "" && password != ""){
                byte[] bytes = Encoding.Default.GetBytes(user_name + ":" + password);
                string auth = "Basic " + Convert.ToBase64String(bytes);
                request.Headers.Add("Authorization", auth);
            }
            request.Method = method;
            request.CookieContainer = msgCookies;
            if (method.ToUpper().Equals("POST"))
            {
                request.ContentType = "application/x-www-form-urlencoded";
                byte[] byteRequest = Encoding.Default.GetBytes(data);
                Stream rs = request.GetRequestStream();
                rs.Write(byteRequest, 0, byteRequest.Length);
                rs.Close();
            }
            try
            {
                HttpWebResponse response = (HttpWebResponse)request.GetResponse();
                HttpStatusCode HSC = response.StatusCode;
                msgCookies.Add(response.Cookies);
                Stream resultStream = response.GetResponseStream();
                StreamReader sr = new StreamReader(resultStream, Encoding.UTF8);
                string html = sr.ReadToEnd();
                responseText = html;
                return true;
            }
            catch (Exception ex)
            {
                responseText = ex.Message;
                return false;
            }
        }
        public static string getCookiesValue(string attribute, string url)
        {
            string value = "";
            Uri uri = new Uri(url);
            CookieCollection cc = msgCookies.GetCookies(uri);
            value = cc[attribute].Value;
            return value;
        }
    }
}
