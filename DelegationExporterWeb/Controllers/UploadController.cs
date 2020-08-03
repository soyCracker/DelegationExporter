using Microsoft.AspNetCore.Mvc;

namespace DelegationExporterWeb.Controllers
{
    public class UploadController : Controller
    {
        public IActionResult Index()
        {
            return View();
        }
    }
}