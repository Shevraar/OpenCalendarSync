namespace OpenCalendarSync.Lib.Location
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

        public override bool Equals(object obj)
        {
            // If parameter is null return false.
            if (obj == null)
            {
                return false;
            }

            // If parameter cannot be cast to Point return false.
            var p = obj as GenericLocation;
            if ((object)p == null)
            {
                return false;
            }

            return  (this.Name == p.Name) &&
                    (this.Latitude == p.Latitude) &&
                    (this.Longitude == p.Longitude);
        }

        public override int GetHashCode()
        {
            return this.Name.GetHashCode();
        }
    }
}