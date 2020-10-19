using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DelegationExporterWeb.Models;
using Microsoft.AspNetCore.Mvc;

namespace DelegationExporterWeb.Controllers
{
    public class TableController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }

        
    }
}
