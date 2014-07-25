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

    public enum ResponseStatus : sbyte
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
    }
}