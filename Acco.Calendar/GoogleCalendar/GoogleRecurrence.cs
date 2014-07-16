using System;
using System.Collections.Generic;

namespace Acco.Calendar.Event
{
    public class GoogleRecurrence : GenericRecurrence
    {
        public override string Get()
        {
            string ret = "";
            string fmt = (TimeZone.CurrentTimeZone.GetUtcOffset(DateTime.Today) < TimeSpan.Zero ? "\\-" : "\\+") + "hh\\:mm";
            string offset = TimeZone.CurrentTimeZone.GetUtcOffset(DateTime.Today).ToString(fmt);
            ret = @"RRULE:FREQ=" + Type.ToString();
            if(Expiry.HasValue)
            {
                ret += ";UNTIL=" + Expiry.Value.ToString("yyyymmddThhmmss") + offset;
            }
            if (Days > DayOfWeek.None)
            {
                ret += ";BYDAY=";
                List<string> days = new List<string>();
                if (Days.HasFlag(DayOfWeek.Monday))
                {
                    days.Add("MO");
                }
                if (Days.HasFlag(DayOfWeek.Tuesday))
                {
                    days.Add("TU");
                }
                if (Days.HasFlag(DayOfWeek.Wednesday))
                {
                    days.Add("WE");
                }
                if (Days.HasFlag(DayOfWeek.Thursday))
                {
                    days.Add("TH");
                }
                if (Days.HasFlag(DayOfWeek.Friday))
                {
                    days.Add("FR");
                }
                if (Days.HasFlag(DayOfWeek.Saturday))
                {
                    days.Add("SA");
                }
                if (Days.HasFlag(DayOfWeek.Sunday))
                {
                    days.Add("SU");
                }
                ret += string.Join(",", days.ToArray()) + ";";
            }
            return ret;
        }

        public override void Parse(string rules)
        {
            string[] details = rules.Split(':');
            foreach (string detail in details[1].Split(';'))
            {
                string attribute = detail.Split('=')[1];
                if (detail.Contains("FREQ"))
                {
                    switch (attribute)
                    {
                        case "DAILY":
                            Type = RecurrencyType.DAILY;
                            break;
                        case "WEEKLY":
                            Type = RecurrencyType.WEEKLY;
                            break;
                        case "MONTHLY":
                            Type = RecurrencyType.MONTHLY;
                            break;
                        case "YEARLY":
                            Type = RecurrencyType.YEARLY;
                            break;
                    }
                }
                else if (detail.Contains("UNTIL"))
                {
                    RFC3339DateTime temporaryExpiryDate = new RFC3339DateTime(attribute);
                }
            }
        }
    }
}
