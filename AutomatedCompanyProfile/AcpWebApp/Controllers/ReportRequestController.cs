using AcpWebApp.Models;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace AcpWebApp.Controllers
{
    public class ReportRequestController : Controller
    {
        // GET: ReportRequest
        public ActionResult Index()
        {
            return View();
        }

        [HttpGet]
        public ActionResult Index(ReportRequest request)
        {
            return View();
        }

    }
}