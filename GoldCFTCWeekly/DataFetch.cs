using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net;
using System.Net.Security;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading.Tasks;

namespace GoldCFTCWeekly
{

    public class DataFetch
    {
        private static string _hisUrl = "https://www.cftc.gov/sites/default/files/files/dea/cotarchives/{0}/futures/other_sf{1}.htm";

        public bool GetGoldCommodity(out List<int> retLst, string soursePage)
        {
            retLst = new List<int>(12);
            int goldOpenInterestCol = 88;
            int goldOtherCol = 90;
            int silverOpenInterestCol = 67;
            int silverOtherCol = 69;

            string[] pageLines = soursePage.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            string strOpenInterest = pageLines[silverOpenInterestCol];
            var openIntLines = System.Text.RegularExpressions.Regex.Split(strOpenInterest, @"\s{2,}");
            strOpenInterest = openIntLines[2];
            strOpenInterest = strOpenInterest.Replace(",", "");
            int openInterest;
            bool ret = Int32.TryParse(strOpenInterest, out openInterest);
            if (!ret)
            {
                return false;
            }
            retLst.Add(openInterest);

            string otherInfo = pageLines[silverOtherCol];
            var dataList = System.Text.RegularExpressions.Regex.Split(otherInfo, @"\s{2,}").ToList();
            dataList.RemoveAt(0);
            for (int index = 0; index < dataList.Count; index++)
            {
                dataList[index] = dataList[index].Replace(",", "");
                dataList[index] = dataList[index].Replace(":", "");
                int tmpValue;
                ret = Int32.TryParse(dataList[index], out tmpValue);
                if (!ret)
                    return false;
                retLst.Add(tmpValue);
            }
            return true;
        }
        public bool GetGoldCommodity(out List<int> retLst, string soursePage, DateTime dt)
        {
            retLst = new List<int>(12);
            int goldOpenInterestCol = 88;
            int goldOtherCol = 90;
            int silverOpenInterestCol = 67;
            int silverOtherCol = 69;

            string[] pageLines = soursePage.Split(Environment.NewLine.ToCharArray(), StringSplitOptions.RemoveEmptyEntries);
            string strOpenInterest = pageLines[silverOpenInterestCol];
            var openIntLines = System.Text.RegularExpressions.Regex.Split(strOpenInterest, @"\s{2,}");
            strOpenInterest = openIntLines[2];
            strOpenInterest = strOpenInterest.Replace(",", "");
            int openInterest;
            bool ret = Int32.TryParse(strOpenInterest, out openInterest);
            if (!ret)
            {
                return false;
            }
            retLst.Add(openInterest);

            string otherInfo = pageLines[silverOtherCol];
            var dataList = System.Text.RegularExpressions.Regex.Split(otherInfo, @"\s{2,}").ToList();
            dataList.RemoveAt(0);
            for (int index = 0; index < dataList.Count; index++)
            {
                dataList[index] = dataList[index].Replace(",", "");
                dataList[index] = dataList[index].Replace(":", "");
                int tmpValue;
                ret = Int32.TryParse(dataList[index], out tmpValue);
                if (!ret)
                    return false;
                retLst.Add(tmpValue);
            }
            return true;
        }

        //public bool FetchData(out List<int> retData)
        //{
        //    retData = new List<int>();
        //    bool isSuccess = false;
        //    WebClient MyWebClient = new WebClient();

        //    MyWebClient.Credentials = CredentialCache.DefaultCredentials;//获取或设置用于对向Internet资源的请求进行身份验证的网络凭据。
        //    string webAddress = @"https://www.cftc.gov/dea/futures/other_sf.htm";
        //    Byte[] pageData = MyWebClient.DownloadData(webAddress);//从指定网站下载数据

        //    //string pageHtml = Encoding.Default.GetString(pageData);  //如果获取网站页面采用的是GB2312，则使用这句             

        //    string pageHtml = Encoding.UTF8.GetString(pageData); //如果获取网站页面采用的是UTF-8，则使用这句
        //    isSuccess = GetGoldCommodity(out retData, pageHtml);

        //    return isSuccess;

        //}

        public bool FetchData(out List<int> retData, ref DateTime date)
        {
            retData = new List<int>();
            bool isSuccess = false;
            if (date.DayOfWeek != DayOfWeek.Tuesday)
                return false;
            int year = date.Year;
            string dateform = date.ToString("MMddyy");
            try
            {
                WebClient myClient = new WebClient();

                myClient.Credentials = CredentialCache.DefaultCredentials;
                //string webAddress = @"https://www.cftc.gov/dea/futures/other_sf.htm";
                string webAddress = string.Format(_hisUrl, year, dateform);
                //myClient.Headers.Add(HttpRequestHeader.ContentType, "text/xml");
                //ServicePointManager.ServerCertificateValidationCallback +=
                //    delegate (object sender, X509Certificate certificate, X509Chain chain,
                //        SslPolicyErrors sslPolicyErrors)
                //    {
                //        return true;
                //    };
                //myClient.Headers[HttpRequestHeader.ContentType] = "application/octet-stream";
                //myClient.Headers[HttpRequestHeader.UserAgent] = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)";
                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                Byte[] pageData = myClient.DownloadData(webAddress); //从指定网站下载数据
                string pageHtml = Encoding.UTF8.GetString(pageData); //如果获取网站页面采用的是UTF-8，则使用这句
                isSuccess = GetGoldCommodity(out retData, pageHtml);
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }

            return isSuccess;
        }

        public bool FetchData(out List<int> retData, ref DateTime date,int offset)
        {
            retData = new List<int>();
            bool isSuccess = false;
            //date = getWeekUpOfDate(date, DayOfWeek.Tuesday, -1);
            DateTime dt = date.AddDays(offset);

            int year = dt.Year;
            string dateform = dt.ToString("MMddyy");
            try
            {
                WebClient myClient = new WebClient();

                myClient.Credentials = CredentialCache.DefaultCredentials;
                //string webAddress = @"https://www.cftc.gov/dea/futures/other_sf.htm";
                string webAddress = string.Format(_hisUrl, year, dateform);
                //myClient.Headers.Add(HttpRequestHeader.ContentType, "text/xml");
                //ServicePointManager.ServerCertificateValidationCallback +=
                //    delegate (object sender, X509Certificate certificate, X509Chain chain,
                //        SslPolicyErrors sslPolicyErrors)
                //    {
                //        return true;
                //    };
                //myClient.Headers[HttpRequestHeader.ContentType] = "application/octet-stream";
                //myClient.Headers[HttpRequestHeader.UserAgent] = "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.2; .NET CLR 1.0.3705;)";
                System.Net.ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls11 | SecurityProtocolType.Tls12;
                Byte[] pageData = myClient.DownloadData(webAddress); //从指定网站下载数据
                string pageHtml = Encoding.UTF8.GetString(pageData); //如果获取网站页面采用的是UTF-8，则使用这句
                isSuccess = GetGoldCommodity(out retData, pageHtml);
                date = dt;
            }
            catch (Exception ex)
            {
                Debug.WriteLine(ex.ToString());
            }

            return isSuccess;
        }

        public DateTime getWeekUpOfDate(DateTime dt, DayOfWeek weekday, int Number)
        {
            int wd1 = (int)weekday;
            int wd2 = (int)dt.DayOfWeek;
            return wd2 == wd1 ? dt.AddDays(7 * Number) : dt.AddDays(7 * Number - wd2 + wd1);
        }

        public bool FetchHistoricData(out List<int> refData, DateTime date)
        {
            refData = new List<int>();
            return false;
        }
    }


}
