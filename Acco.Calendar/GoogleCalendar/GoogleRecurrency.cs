using System;

namespace Acco.Calendar.Event
{
    public class GoogleRecurrency : GenericRecurrency
    {
        public override string ToString()
        {
            string ret = "";
            string offset = TimeZone.CurrentTimeZone.GetUtcOffset(DateTime.Today).ToString("-hh:mm");
            // yyyymmddThhmmss-OFFSET
            // 20110701T100000-07:00
            ret = @"RRULE:FREQ=" + Type.ToString() + ";UNTIL=" + Expiry.Value.ToString("yyyymmddThhmmss") + offset;
            return ret;
        }

        public void FromString(string rules)
        {
            // guarda un po' cosa viene fuori...
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
