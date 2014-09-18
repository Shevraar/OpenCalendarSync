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

        public static bool operator ==(GenericLocation l1, GenericLocation l2)
        {
            if((object)l1 != null && (object)l2 != null)
            {
                return  (l1.Name == l2.Name) &&
                        (l1.Latitude == l2.Latitude) &&
                        (l1.Longitude == l2.Longitude);
            }
            else
            {
                return (object)l1 == (object)l2;
            }   
        }

        public static bool operator !=(GenericLocation l1, GenericLocation l2)
        {
            return !(l1 == l2);
        }

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

            return (this == p);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}