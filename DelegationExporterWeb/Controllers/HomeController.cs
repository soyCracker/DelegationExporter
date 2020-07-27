using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using DelegationExporterWeb.Models;
using Microsoft.AspNetCore.Http;
using System.IO;
using Microsoft.AspNetCore.Hosting;

namespace DelegationExporterWeb.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly string _folder;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment env)
        {
            _logger = logger;
            // 把上傳目錄設為：wwwroot\UploadFolder
            _folder = $@"{env.WebRootPath}";
        }

        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        [HttpPost]
        public async Task<IActionResult> Upload(List<IFormFile> files)
        {
            var size = files.Sum(f => f.Length);

            foreach (var file in files)
            {
                if (file.Length > 0)
                {
                    /*var path = $@"{_folder}\{file.FileName}";
                    using (var stream = new FileStream(path, FileMode.Create))
                    {
                        await file.CopyToAsync(stream).ConfigureAwait(false);
                    }*/

                }
            }

            return Ok(new { count = files.Count, size });
        }
    }
}
