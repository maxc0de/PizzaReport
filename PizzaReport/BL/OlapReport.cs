using System;
using System.IO;
using System.Net;
using System.Security.Cryptography;
using System.Text;

namespace MilanoExtraReport.BL
{
    public class OlapReport
    {
        private static string login = "admin";
        private static string password = "Zx08365#";
        private static string address = "31.173.34.137";
        private static string port = "19080";

        public string GetReport()
        {
            string key = Login();

            string report;
            try
            {
                report = GetReport(key);
            }
            finally
            {
                LogOut(key);
            }

            return report;
        }

        private string Login()
        {
            return GET($"http://{address}:{port}/resto/api/auth", $"login={login}&pass={GetHash(password)}");
        }

        private void LogOut(string key)
        {
            GET($"http://{address}:{port}/resto/api/logout", $"key={key}");
        }

        private static string GetReport(string key)
        {
            return GET($"http://{address}:{port}/resto/api/reports/olap", $"key={key}&report=SALES&from={"30.04.2020"}&to={"01.05.2020"}&groupRow=Department&groupRow=DishGroup&groupRow=DishName&agr=DishAmountInt&agr=fullSum");
        }

        private static string GetHash(string input)
        {
            using (SHA1Managed sha1 = new SHA1Managed())
            {
                var hash = sha1.ComputeHash(Encoding.Default.GetBytes(input));
                var sb = new StringBuilder(hash.Length * 2);

                foreach (byte b in hash)
                {
                    sb.Append(b.ToString("x2"));
                }

                return sb.ToString();
            }
        }

        private static string GET(string url, string data)
        {
            WebRequest req = WebRequest.Create(url + "?" + data);
            WebResponse resp = req.GetResponse();
            Stream stream = resp.GetResponseStream();
            StreamReader sr = new StreamReader(stream);
            string Out = sr.ReadToEnd();
            sr.Close();
            return Out;
        }
    }
}
