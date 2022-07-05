namespace OpenXmlBinder.Sample
{
    public class WeatherForecast
    {
        public DateTime Date { get; set; }

        public int TemperatureC { get; set; }

        public int TemperatureF => 32 + (int)(TemperatureC / 0.5556);

        public Location Location { get; set; } = null!;

        public string? Summary { get; set; }

        public WeatherForecast(DateTime date, int temperatureC, Location location, string? summary)
        {
            Date = date;
            TemperatureC = temperatureC;
            Location = location;
            Summary = summary;
        }

        public WeatherForecast()
        {

        }
    }
}