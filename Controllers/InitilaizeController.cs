using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;

namespace BrandixAutomation.Labdip.API.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class InitilaizeController : ControllerBase
    {

        private readonly ILogger<InitilaizeController> _logger;

        public InitilaizeController(ILogger<InitilaizeController> logger)
        {
            _logger = logger;
        }

        [HttpGet]
        public String Get()
        {
            return "Initilaized :)";
        }
    }
}
