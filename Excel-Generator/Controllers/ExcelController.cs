using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;

namespace Excel_Generator.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class ExcelController : ControllerBase
    {
        private readonly ILogger<ExcelController> _logger;

        public ExcelController(ILogger<ExcelController> logger)
        {
            _logger = logger;
        }    
    }
}
