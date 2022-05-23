using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Diagnostics;
using System.Linq;
using System.Net.Sockets;
using System.Web;
using System.Web.Mvc;

namespace AcpWebApp.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
        }

        public string GenerateReport()
        {
            var rawUrl = Request.RawUrl;
            rawUrl = HttpUtility.UrlDecode(rawUrl);
            int q = rawUrl.IndexOf('?');
            var progress = SendToAcpServer("GENERATEREPORT!" + rawUrl.Substring(q + 1));
            return HandleReportProgressResponse(progress);
        }

        public string HandleReportProgressResponse(string data)
        {
            return data;
            /*
            if (reportFilename != null)
            {
                var slashed = reportFilename.Replace('\\', '/');
                var i = slashed.IndexOf("/Output/", StringComparison.CurrentCultureIgnoreCase);
                if (i < 0)
                    return null;
                var fn = slashed.Substring(i);
                return fn;
            }
            return null;
            */
        }

        public string GetReportProgress()
        {
            var rawUrl = Request.RawUrl;
            rawUrl = HttpUtility.UrlDecode(rawUrl);
            int q = rawUrl.IndexOf('?');
            var progress = SendToAcpServer("GETREPORTPROGRESS!" + rawUrl.Substring(q + 1));
            return HandleReportProgressResponse(progress);
        }


        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";
            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";
            return View();
        }

        [HttpGet]
        public ActionResult SearchByName()
        {
            var rawUrl = Request.RawUrl;
            rawUrl = HttpUtility.UrlDecode(rawUrl);
            int q = rawUrl.IndexOf('?');
            var response = SendToAcpServer("SEARCHBYNAME!" + rawUrl.Substring(q + 1));
            return Content(response);
        }

        [HttpGet]
        public ActionResult FindPeers()
        {
            var rawUrl = Request.RawUrl;
            rawUrl = HttpUtility.UrlDecode(rawUrl);
            int q = rawUrl.IndexOf('?');
            var response = SendToAcpServer("FINDPEERS!" + rawUrl.Substring(q + 1));
            return Content(response);
        }

        [HttpGet]
        public ActionResult QuickTickerSearch()
        {
            var rawUrl = Request.RawUrl;
            rawUrl = HttpUtility.UrlDecode(rawUrl);
            int q = rawUrl.IndexOf('?');
            var response = SendToAcpServer("QUICKTICKERSEARCH!" + rawUrl.Substring(q + 1));
            return Content(response);
        }

        [HttpGet]
        public ActionResult ValidateTimePeriodSettings()
        {
            var rawUrl = Request.RawUrl;
            rawUrl = HttpUtility.UrlDecode(rawUrl);
            int q = rawUrl.IndexOf('?');
            var response = SendToAcpServer("VALIDATETIMEPERIODSETTINGS!" + rawUrl.Substring(q + 1));
            return Content(response);
        }

        [HttpGet]
        public ActionResult Unload()
        {
            var rawUrl = Request.RawUrl;
            rawUrl = HttpUtility.UrlDecode(rawUrl);
            int q = rawUrl.IndexOf('?');
            SendToAcpServer("QUIT!" + rawUrl.Substring(q + 1));
            return null;
        }

        public string SendToAcpServer(string request)
        {
            var index = request.IndexOf("Sender=");
            if (index < 0)
                return "ERROR=Malformed request";
            index += 7;
            var end = request.IndexOf("&", index);
            var sender = (end < 0) ? request.Substring(index) : request.Substring(index, end - index);

            var client = OpenConnectionToAcpServer(sender);
            if (client == null)
                return "ERROR=The ACP server is not available";

            var bytes = System.Text.Encoding.UTF8.GetBytes(request);
            var stream = client.GetStream();
            stream.Write(bytes, 0, bytes.Length);

            bytes = new byte[4096];
            int length;
            try
            {
                length = client.Connected ? stream.Read(bytes, 0, bytes.Length) : 0;
            }
            catch
            {
                length = 0;
            }

            var response = (length > 0) ? System.Text.Encoding.UTF8.GetString(bytes, 0, length) : "";
            if (response == " ")
                response = "";

            stream.Close();
            client.Close();
            return response;
        }

        private TcpClient OpenConnectionToAcpServer(string sender)
        {
            var port = 1235;
            try
            {
                var client = new TcpClient("localhost", port);
                client.SendTimeout = 15 * 60 * 1000;
                client.ReceiveTimeout = 15 * 60 * 1000;
                Session[sender + "TcpClient"] = client;
                return client;
            }
            catch
            {
                Session[sender + "TcpClient"] = null;
                return null;
            }
        }

        private void ShutConnectionToAcpServer(string sender, ref TcpClient client)
        {
            if ((client != null) && !client.Connected)
            {
                NetworkStream stream = null;
                try { stream = client.GetStream(); } catch { }
                if (stream != null)
                {
                    var bytes = System.Text.Encoding.UTF8.GetBytes("QUIT!");
                    try { stream.Write(bytes, 0, bytes.Length); } catch { }
                    stream.Close();
                }
                Session[sender + "TcpClient"] = null;
                client.Close();
                client = null;
            }
        }

        [Route("Home/Output/{folder}/{filename}.pptx")]
        public ActionResult Output(string folder, string filename)
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}