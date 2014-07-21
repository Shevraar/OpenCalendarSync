namespace Acco.Calendar.Location
{
    public interface ILocation
    {
        string Name { get; set; }

        int? Longitude { get; set; }

        int? Latitude { get; set; }
    }

    public class GenericLocation : ILocation
    {
        public string Name { get; set; }

        public int? Longitude { get; set; }

        public int? Latitude { get; set; }
    }
}