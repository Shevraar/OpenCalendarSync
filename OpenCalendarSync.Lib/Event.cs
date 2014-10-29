using System.Linq;
using MongoDB.Bson.Serialization.Attributes;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using OpenCalendarSync.Lib.Location;
using OpenCalendarSync.Lib.Person;


namespace OpenCalendarSync.Lib.Event
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

        [BsonIgnore]
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
            // wip - move this somewhere else, or use an ordered list to insert attendees in a ordered manner
            this.Attendees.Sort((a1, a2) => String.Compare(a1.Email, a2.Email, StringComparison.Ordinal));
            p.Attendees.Sort((a1, a2) => String.Compare(a1.Email, a2.Email, StringComparison.Ordinal));
            var idIsEqual = this.Id == p.Id;
            var startDateIsEqual = this.Start == p.Start;
            var endDateIsEqual = this.End == p.End;
            var locationIsEqual = this.Location.Equals(p.Location);
            var descriptionIsEqual = this.Description == p.Description;
            var recurrenceIsEqual = true;
            if(this.Recurrence != null && p.Recurrence != null)
                recurrenceIsEqual = this.Recurrence.Pattern == p.Recurrence.Pattern;
            var attendeesCountIsEqual = this.Attendees.Count == p.Attendees.Count /* first check if the number of attendees is the same */;
            var attendeesAreEqual = !this.Attendees.Except(p.Attendees).Any();
            //
            return  (idIsEqual) &&
                    (startDateIsEqual) &&
                    (endDateIsEqual) &&
                    (locationIsEqual) &&
                    (descriptionIsEqual) &&
                    (recurrenceIsEqual) &&
                    (attendeesCountIsEqual) &&
                    (attendeesAreEqual);
        }

        public override int GetHashCode()
        {
            return this.Id.GetHashCode();
        }
    }
}