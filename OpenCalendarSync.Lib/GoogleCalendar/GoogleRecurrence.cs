//

using System;
using System.Collections.Generic;
using RecPatt = DDay.iCal.RecurrencePattern;

//
namespace OpenCalendarSync.Lib.Event
{
    public class GoogleRecurrence : GenericRecurrence
    {
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        //public override string Get()
        //{
        //    var modifiedPattern = RecurrencePattern.ToString();
        //    // todo: modify this pattern to add timezones manually...
        //    modifiedPattern = "RRULE:" + modifiedPattern;
        //    return modifiedPattern;
        //}

        public override void Parse<T>(T rules)
        {
            if (rules is List<string>)
            {
                Log.Info(String.Format("Parsing GoogleRecurrence [{0}]", rules));
                Pattern = rules as List<string>;
                Log.Debug(String.Format("Recurrence pattern is [{0}]", Pattern));
            }
            else
            {
                throw new RecurrenceParseException("GoogleRecurrence: Unsupported type.", typeof(T));
            }
        }
    }
}