using System;
using System.ComponentModel.DataAnnotations;
using OpenCalendarSync.Lib.Utilities;
using OutlookResponseStatusEnum = Microsoft.Office.Interop.Outlook.OlResponseStatus;

namespace OpenCalendarSync.Lib.Person
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

            return  (this.Email == p.Email) &&
                    (this.Name == p.Name) &&
                    (this.FirstName == p.FirstName) &&
                    (this.LastName == p.LastName);
        }

        public override int GetHashCode()
        {
            return this.Email.GetHashCode();
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

    /// <summary>
    /// Specialization of GenericPerson, with ResponseStatus as an additional parameter
    /// </summary>
    public class GenericAttendee : GenericPerson
    {
        /// <summary>
        /// The response status of the meeting/event/appointment attendee
        /// </summary>
        public ResponseStatus Response { get; set; }

        public override bool Equals(object obj)
        {
            // If parameter is null return false.
            if (obj == null)
            {
                return false;
            }

            // If parameter cannot be cast to Point return false.
            var p = obj as GenericAttendee;
            if ((object)p == null)
            {
                return false;
            }

            return (base.Equals(p)) &&
                   (this.Response == p.Response);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }
    }
}