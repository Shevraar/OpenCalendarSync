using System;
using System.Collections.Generic;
using MongoDB.Bson.Serialization;

using RecPatt = DDay.iCal.RecurrencePattern;

namespace OpenCalendarSync.Lib.Event
{
    public interface IRecurrence
    {
        void Parse<T>(T rules);
    }

    /// <summary>
    /// This is the base class for defining Recurrence behaviour
    /// </summary>
    public class GenericRecurrence : IRecurrence
    {
        protected RecPatt RecurrencePattern { get; set; }
        protected string Exdate { get; set; }

        protected GenericRecurrence()
        {
            if (!BsonClassMap.IsClassMapRegistered(typeof(RecPatt)))
            {
                BsonClassMap.RegisterClassMap<RecPatt>();
            }
        }

        public virtual void Parse<T>(T rules)
        {
            throw new NotImplementedException();
        }

        public List<string> Pattern
        {
            get
            {
                return new List<string>
                {
                    "RRULE:" + RecurrencePattern,
                    Exdate
                };
            }
            set
            {
                RecurrencePattern = new RecPatt(value[0]);
                Exdate = value[1];
            }
        }
    }

    [Serializable]
    public class RecurrenceParseException : Exception
    {
        /// <summary>
        /// Provide a message and the type of recurrence throwing the exception
        /// </summary>
        /// <param name="message"></param>
        /// <param name="typeOfRule"></param>
        public RecurrenceParseException(string message, Type typeOfRule) :
            base(message)
        {
            TypeOfRule = typeOfRule;
        }

        public Type TypeOfRule { get; private set; }
    }
}