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
        private NecessaryFileService necessaryFileService;

        public FileController(DelegationExporterDBContext context)
        {
            excelService = new ExcelService();
            pdfServce = new PdfServce(context);
            necessaryFileService = new NecessaryFileService(context);
        }

        [HttpPost("GetDelegation")]
        [HttpPost]
        public IActionResult GetDelegation(IFormFile file)
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
                necessaryFileService.SaveFileToDB(file);
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

        [HttpPost("GetData")]
        [HttpPost]
        public IActionResult GetData()
        {
            List<TableModel> list = new List<TableModel>();
            TableModel table1 = new TableModel
            {
                Name = "Lai",
                Age = 25,
                Money = 100
            };

            TableModel table2 = new TableModel
            {
                Name = "Fen",
                Age = 30,
                Money = 120
            };

            TableModel table3 = new TableModel
            {
                Name = "Yu",
                Age = 27,
                Money = 95
            };

            list.Add(table1);
            list.Add(table2);
            list.Add(table3);

            return Ok(new { Value = true, Response = list });
        }
    }
}