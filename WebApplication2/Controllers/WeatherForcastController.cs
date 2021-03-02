using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace WebApplication2.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class WeatherForcast : ControllerBase
    {
        //private static readonly string[] Summaries = new[]
        //{
        //    "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
        //};

        //private readonly ILogger<ReceivedInfosController> _logger;

        //public ReceivedInfosController(ILogger<ReceivedInfosController> logger)
        //{
        //    _logger = logger;
        //}

        [HttpGet]
        public string  Get()
        {
            return "";
            var rng = new Random();
            //return Enumerable.Range(1, 5).Select(index => new ReceivedInfos
            //{
            //    timeStamp = 1
            //})
            //.ToArray();
        }
    }
}
