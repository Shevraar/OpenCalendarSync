using System.Linq;
using Acco.Calendar.Location;
using Acco.Calendar.Person;
using MongoDB.Bson.Serialization;
using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;


namespace Acco.Calendar.Event
{
    public interface IEvent
    {
        [Required(ErrorMessage = "This field is required")]
        string Id { get; set; }

        GenericPerson Organizer { get; set; }

        GenericPerson Creator { get; set; }

        DateTime? Created { get; set; }

        DateTime? LastModified { get; set; }

        DateTime? Start { get; set; }

        DateTime? End { get; set; }

        [Required(ErrorMessage = "This field is required")]
        string Summary { get; set; }

        [Required(ErrorMessage = "This field is required")]
        string Description { get; set; }

        [Required(ErrorMessage = "This field is required")]
        GenericLocation Location { get; set; }

        GenericRecurrence Recurrence { get; set; }

        List<GenericAttendee> Attendees { get; set; }

        EventAction Action { get; set; }
    }

    public enum EventAction : sbyte
    {
        Add = 0,
        Update,
        Remove,
        Duplicate
    }

    public class GenericEvent : IEvent
    {
        public GenericEvent(string id)
        {
            Id = id;
            Summary = "No Summary";
            Description = "No Description";
            Location = new GenericLocation { Name = "No Location" };
        }

        public GenericEvent(string id, string summary, string description)
        {
            Id = id;
            Summary = summary;
            Description = description;
            Location = new GenericLocation { Name = "No Location" };
        }

        public GenericEvent(string id, string summary, string description, ILocation location)
        {
            Id = id;
            Summary = summary;
            Description = description;
            Location = location as GenericLocation;
        }

        public string Id { get; set; }

        public GenericPerson Organizer { get; set; }

        public GenericPerson Creator { get; set; }

        [BsonDateTimeOptions(Kind = DateTimeKind.Local)]
        public DateTime? Created { get; set; }

        [BsonDateTimeOptions(Kind = DateTimeKind.Local)]
        public DateTime? LastModified { get; set; }

        [BsonDateTimeOptions(Kind = DateTimeKind.Local)]
        public DateTime? Start { get; set; }

        [BsonDateTimeOptions(Kind = DateTimeKind.Local)]
        public DateTime? End { get; set; }

        public string Summary { get; set; }

        public string Description { get; set; }

        public GenericLocation Location { get; set; }

        public GenericRecurrence Recurrence { get; set; }

        public List<GenericAttendee> Attendees { get; set; }

        public override string ToString()
        {
            var eventString = "[";
            eventString += "Id: " + Id;
            eventString += "\n";
            eventString += "Summary: " + Summary;
            eventString += "\n";
            eventString += "Location: " + Location.Name;
            eventString += "]";
            return eventString;
        }

        public EventAction Action { get; set; }

        public override bool Equals(object obj)
        {
            // If parameter is null return false.
            if (obj == null)
            {
                return false;
            }

            // If parameter cannot be cast to Point return false.
            var p = obj as GenericEvent;
            if ((object)p == null)
            {
                return false;
            }

            // Event comparison - Id, Dates, Location, Recurrence, Attendees
            return (this == p);
        }

        public override int GetHashCode()
        {
            return base.GetHashCode();
        }

        public static bool operator ==(GenericEvent e1, GenericEvent e2)
        {
            if ((object)e1 != null &&
                (object)e2 != null)
            {
                // wip - move this somewhere else, or use an ordered list to insert attendees in a ordered manner
                e1.Attendees.Sort((a1, a2) => String.Compare(a1.Email, a2.Email, StringComparison.Ordinal));
                e2.Attendees.Sort((a1, a2) => String.Compare(a1.Email, a2.Email, StringComparison.Ordinal));
                //
                return  (e1.Id == e2.Id) &&
                        (e1.Start == e2.Start) &&
                        (e1.End == e2.End) &&
                        (e1.Location == e2.Location) &&
                        (e1.Description == e2.Description) &&
                        (e1.Recurrence != null && e2.Recurrence != null) &&
                        (e1.Recurrence.Pattern == e2.Recurrence.Pattern) &&
                        (e1.Attendees.Count == e2.Attendees.Count /* first check if the number of attendees is the same */) &&
                        (!e1.Attendees.Except(e2.Attendees).Any());
            }
            return (object)e1 == (object)e2;
        }

        public static bool operator !=(GenericEvent e1, GenericEvent e2)
        {
            return !(e1 == e2);
        }
    }
}