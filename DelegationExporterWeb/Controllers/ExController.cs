using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;

namespace DelegationExporterWeb.Controllers
{
    public class ExController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Clock()
        {
            return View();
        }

        public IActionResult Table()
        {
            return View();
        }
    }
}
