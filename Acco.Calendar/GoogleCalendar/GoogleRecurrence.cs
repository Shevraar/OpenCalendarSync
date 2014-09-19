//

using System;
using RecPatt = DDay.iCal.RecurrencePattern;

//
namespace Acco.Calendar.Event
{
    public class GoogleRecurrence : GenericRecurrence
    {
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public override string Get()
        {
            var modifiedPattern = _RecPatt.ToString();
            // todo: modify this pattern to add timezones manually...
            modifiedPattern = "RRULE:" + modifiedPattern;
            return modifiedPattern;
        }

        public override void Parse<T>(T rules)
        {
            if (rules is string)
            {
                Log.Info(String.Format("Parsing GoogleRecurrence [{0}]", rules));
                var stringRules = rules as string;
                _RecPatt = new RecPatt(stringRules);
                Log.Debug(String.Format("iCalendar recurrence pattern is [{0}]", _RecPatt));
            }
            else
            {
                throw new RecurrenceParseException("GoogleRecurrence: Unsupported type.", typeof(T));
            }
        }
    }
}