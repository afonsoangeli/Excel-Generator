using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Pecege.MoveHumaniza.Domain.Extensions;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace Excel_Generator.Controllers
{
    [ApiController]
    [Route("api/excels")]
    public class ExcelController : ControllerBase
    {
        [HttpGet]
        [ProducesResponseType(StatusCodes.Status200OK)]
        public async Task<IActionResult> Download()
        {
            try
            {
                List<string[]> data = new() {
                    new string[] { "Name", "Age", "Email" }, // Header columns data
                };

                List<string[]> dataFromDataBase = new() // Here is where you get your data from your database
                {
                    new string[] { "Contoso", "21", "contoso@domain.com" },
                    new string[] { "Foo", "48", "foo@domain.com" },
                    new string[] { "FooTwo", "23", "footwo@domain.com" },
                };

                data.AddRange(dataFromDataBase);

                var excel = new Excel("ContosoSheet", "This is my awesome sheet", data);
                var memoryStream = excel.GenerateExcelFile();
                memoryStream.Seek(0, SeekOrigin.Begin);

                var file = File(memoryStream, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", $"contoso.xlsx");

                return file;
            }
            catch (Exception e)
            {
                return Ok(e);
            }
        }
    }
}
