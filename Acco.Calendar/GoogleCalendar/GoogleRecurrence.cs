//
using DDay.iCal;

//
namespace Acco.Calendar.Event
{
    public class GoogleRecurrence : GenericRecurrence
    {
        public override string Get()
        {
            var modifiedPattern = Pattern.ToString();
            // todo: modify this pattern to add timezones manually...
            modifiedPattern = "RRULE:" + modifiedPattern;
            return modifiedPattern;
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