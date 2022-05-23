using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Sockets;
using System.Web;
using System.Web.Mvc;
using System.Web.Optimization;
using System.Web.Routing;

namespace AcpWebApp
{
    public class MvcApplication : System.Web.HttpApplication
    {
        protected void Application_Start()
        {
            AreaRegistration.RegisterAllAreas();
            FilterConfig.RegisterGlobalFilters(GlobalFilters.Filters);
            RouteConfig.RegisterRoutes(RouteTable.Routes);
            BundleConfig.RegisterBundles(BundleTable.Bundles);
        }

        protected void Session_Start(Object sender, EventArgs e)
        {
            Session.Add("StartTime", DateTime.Now);

            /*
            try
            {
                var port = 1235;
                var client = new TcpClient("localhost", port);
                client.SendTimeout = 15 * 60 * 1000;
                client.ReceiveTimeout = 15 * 60 * 1000;
                Session["TcpClient"] = client;
            }
            catch
            {
                Session["TcpClient"] = null;
            }
            */
        }

        protected void Session_End(Object sender, EventArgs e)
        {
            /*
            var client = (TcpClient) Session["TcpClient"];
            var stream = client?.GetStream();

            if (stream != null)
            {
                var bytes = System.Text.Encoding.UTF8.GetBytes("QUIT!");
                try { stream.Write(bytes, 0, bytes.Length); } catch { }
                stream.Close();
            }

            if (client != null)
            {
                Session["TcpClient"] = null;
                client.Close();
            }
            */
        }
    }
}
