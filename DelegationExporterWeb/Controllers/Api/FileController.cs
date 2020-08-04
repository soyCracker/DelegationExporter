using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using DelegationExporterEntity.Entities;
using DelegationExporterWeb.Models;
using DelegationExporterWeb.Service;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;

namespace DelegationExporterWeb.Controllers.Api
{
    [Route("api/[controller]")]
    [ApiController]
    public class FileController : ControllerBase
    {
        private ExcelService excelService;
        private PdfServce pdfServce;

        public FileController(DelegationExporterDBContext context)
        {
            excelService = new ExcelService();
            pdfServce = new PdfServce(context);
        }

        [HttpPost("UploadXls")]
        [HttpPost]
        public IActionResult UploadXls(IFormFile file)
        {
            if (file.Length > 0)
            {
                List<DelegationModel> delegationList = excelService.ReadDelegation(file);
                pdfServce.WriteInDelegation(delegationList);
                return Ok(new { Value = true, ErrorCode = delegationList.Count });
            }
            return Ok(new { Value = false, ErrorCode = 856 });
        }

        /*[HttpPost("UploadFile")]
        [HttpPost]
        public IActionResult UploadFile(IFormFile file)
        {
            if (file.Length > 0)
            {
                pdfServce.SaveFileToDB(file);
                return Ok(new { Value = true, ErrorCode = 0 });
            }
            return Ok(new { Value = false, ErrorCode = 856 });
        }*/

        [HttpPost("Test")]
        [HttpPost]
        public IActionResult Test()
        {
            return Ok(new { Value = true });
        }
    }
}