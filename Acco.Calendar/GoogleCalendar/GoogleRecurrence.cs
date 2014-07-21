using System;
using System.Collections.Generic;
//
using DDay.iCal;
//
namespace Acco.Calendar.Event
{
    public class GoogleRecurrence : GenericRecurrence
    {
        public GoogleRecurrence()
        {
            
        }
        public override string Get()
        {
            var _modifiedPattern = Pattern.ToString();
            // todo: modify this pattern to add timezones manually...
            _modifiedPattern = "RRULE:" + _modifiedPattern;
            return _modifiedPattern;
        }

        public override void Parse<T>(T rules)
        {
            if (rules is string)
            {
                var stringRules = rules as string;
                Pattern = new RecurrencePattern(stringRules);
            }
            else
            {
                throw new RecurrenceParseException("GoogleRecurrence: Unsupported type.", typeof(T));
            }
        }
    }
}
