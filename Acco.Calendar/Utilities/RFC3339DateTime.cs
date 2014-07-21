using System;

namespace Acco.Calendar.Utilities
{
    public struct Rfc3339DateTime : IEquatable<Rfc3339DateTime>, IComparable<Rfc3339DateTime>
    {
        private readonly DateTimeOffset _value;

        private static readonly string[] Formats = { "yyyy-MM-ddTHH:mm:ssK", "yyyy-MM-ddTHH:mm:ss.ffK", "yyyy-MM-ddTHH:mm:ssZ", "yyyy-MM-ddTHH:mm:ss.ffZ" };

        public Rfc3339DateTime(string rfc3339FormattedDateTime)
        {
            DateTimeOffset tmp;
            if (!DateTimeOffset.TryParseExact(rfc3339FormattedDateTime, Formats, System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.AssumeUniversal, out tmp))
            {
                throw new ArgumentException("Value is not in proper RFC3339 format", "rfc3339FormattedDateTime");
            }
            _value = tmp;
        }

        public static explicit operator Rfc3339DateTime(string rfc3339FormattedDateTime)
        {
            return new Rfc3339DateTime(rfc3339FormattedDateTime);
        }

        public override int GetHashCode()
        {
            return _value.GetHashCode();
        }

        public override bool Equals(object obj)
        {
            if (obj == null) return false;
            if (!(obj is Rfc3339DateTime)) return false;
            return _value.Equals(((Rfc3339DateTime)obj)._value);
        }

        public bool Equals(Rfc3339DateTime other)
        {
            return _value.Equals(other._value);
        }

        public static bool operator ==(Rfc3339DateTime a, Rfc3339DateTime b)
        {
            return a._value == b._value;
        }

        public static bool operator !=(Rfc3339DateTime a, Rfc3339DateTime b)
        {
            return a._value != b._value;
        }

        public int CompareTo(Rfc3339DateTime other)
        {
            return _value.CompareTo(other._value);
        }
    }
}