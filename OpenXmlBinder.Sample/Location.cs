namespace OpenXmlBinder.Sample
{
    public class Location
    {
        public string City { get; set; }
        public string Country { get; set; }

        public Location(string city, string country)
        {
            City = city;
            Country = country;
        }
    }
}
