using Acco.Calendar.Utilities;
using System;
using System.ComponentModel.DataAnnotations;
using OutlookResponseStatusEnum = Microsoft.Office.Interop.Outlook.OlResponseStatus;

namespace Acco.Calendar.Person
{
    public interface IPerson
    {
        string Email { get; set; }

        string Name { get; set; }

        string FirstName { get; set; }

        string LastName { get; set; }
    }

    public class GenericPerson : IPerson
    {
        [RegularExpression(Defines.EmailRegularExpression,
                            ErrorMessage = "Not a valid email address")]
        public string Email { get; set; }

        public string Name { get; set; }

        public string FirstName { get; set; }

        public string LastName { get; set; }

        public static bool operator ==(GenericPerson p1, GenericPerson p2)
        {
            if((object)p1 != null && (object)p2 != null)
            {
                return  (p1.Email == p2.Email) &&
                        (p1.Name == p2.Email) &&
                        (p1.FirstName == p2.FirstName) &&
                        (p1.LastName == p2.LastName);
            }
            return (object)p1 == (object)p2;
        }

        public static bool operator !=(GenericPerson p1, GenericPerson p2)
        {
            return !(p1 == p2);
        }

        public override bool Equals(object obj)
        {
            // If parameter is null return false.
            if (obj == null)
            {
                return false;
            }

            // If parameter cannot be cast to Point return false.
            var p = obj as GenericPerson;
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

    class GoogleResponseStatus : Attribute
    {
        public string Text { get; private set; }

        public GoogleResponseStatus(string text)
        {
            Text = text;
        }
    }

    class OutlookResponseStatus : Attribute
    {
        public Microsoft.Office.Interop.Outlook.OlResponseStatus OlResponse { get; private set; }

        public OutlookResponseStatus(Microsoft.Office.Interop.Outlook.OlResponseStatus olResponse)
        {
            OlResponse = olResponse;
        }
    }

    public enum ResponseStatus : ushort
    {
        [GoogleResponseStatus("accepted"),
        OutlookResponseStatus(OutlookResponseStatusEnum.olResponseAccepted)]
        Accepted = 0,

        [GoogleResponseStatus("declined"),
        OutlookResponseStatus(OutlookResponseStatusEnum.olResponseDeclined)]
        Declined,

        [GoogleResponseStatus("tentative"),
        OutlookResponseStatus(OutlookResponseStatusEnum.olResponseTentative)]
        Tentative,

        [GoogleResponseStatus("accepted"),
        OutlookResponseStatus(OutlookResponseStatusEnum.olResponseOrganized)]
        Organized,

        [GoogleResponseStatus("needsAction"),
        OutlookResponseStatus(OutlookResponseStatusEnum.olResponseNone)]
        None,

        [GoogleResponseStatus("needsAction"),
        OutlookResponseStatus(OutlookResponseStatusEnum.olResponseNotResponded)]
        NotResponded
    }

    public class GenericAttendee : GenericPerson
    {
        public ResponseStatus Response { get; set; }

        public override bool Equals(object obj)
        {
            // If parameter is null return false.
            if (obj == null)
            {
                return false;
            }

            // If parameter cannot be cast to Point return false.
            var p = obj as GenericPerson;
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

        public static bool operator == (GenericAttendee a1, GenericAttendee a2)
        {
            if ((object)a1 != null && (object)a2 != null)
            {
                return a1.Response == a2.Response;
            }
            return (object)a1 == (object)a2;
        }

        public static bool operator !=(GenericAttendee a1, GenericAttendee a2)
        {
            return !(a1 == a2);
        }
    }
}