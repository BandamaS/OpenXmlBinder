using Microsoft.AspNetCore.Mvc;

namespace OpenXmlBinder.Sample.Controllers
{
    [ApiController]
    [Route("api/[controller]/[action]")]
    public class WeatherForecastController : ControllerBase
    {
        private static readonly string[] Summaries = new[]
        {
        "Freezing", "Bracing", "Chilly", "Cool", "Mild", "Warm", "Balmy", "Hot", "Sweltering", "Scorching"
    };

        private readonly ILogger<WeatherForecastController> _logger;

        public WeatherForecastController(ILogger<WeatherForecastController> logger)
        {
            _logger = logger;
        }

        [HttpGet(Name = "GetWeatherForecast")]
        public IEnumerable<WeatherForecast> Get()
        {
            return Enumerable.Range(1, 5).Select(index => new WeatherForecast
            {
                Date = DateTime.Now.AddDays(index),
                TemperatureC = Random.Shared.Next(-20, 55),
                Summary = Summaries[Random.Shared.Next(Summaries.Length)]
            })
            .ToArray();
        }

        [HttpGet(Name = "GetReport")]
        public IActionResult GetReport()
        {
            OpenXmlBinder binder = new OpenXmlBinder("Template/WeatherForecastTemplate.docx");

            binder.AddVariable(new WeatherForecast
            {
                Date = DateTime.Now,
                TemperatureC = Random.Shared.Next(-20, 55),
                Summary = Summaries[Random.Shared.Next(Summaries.Length)],
                Location=new Location("Paris","France")
            });
            binder.AddVariable("Date_Short",DateTime.Now.ToString("d"));

            byte[] package = binder.Generate();

            return File(package, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "WeatherForecast.docx");
        }

    }
}