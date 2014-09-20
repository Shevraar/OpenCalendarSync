using System;
using MongoDB.Bson.Serialization;
using RecPatt = DDay.iCal.RecurrencePattern;

namespace Acco.Calendar.Event
{
    public interface IRecurrence
    {
        void Parse<T>(T rules);

        string Get();
    }

    public class GenericRecurrence : IRecurrence
    {
        protected RecPatt _RecPatt { get; set; }

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

        public virtual string Get()
        {
            throw new NotImplementedException();
        }

        public override string ToString()
        {
            var s = "";
            if (_RecPatt != null)
            {
                s = _RecPatt.ToString();
            }
            return s;
        }

        public string Pattern
        {
            get
            {
                return _RecPatt.ToString();
            }
            set
            {
                _RecPatt = new RecPatt(value);
            }
        }
    }

    [Serializable]
    public class RecurrenceParseException : Exception
    {
        public RecurrenceParseException(string message, Type typeOfRule) :
            base(message)
        {
            TypeOfRule = typeOfRule;
        }

        public Type TypeOfRule { get; private set; }
    }
}