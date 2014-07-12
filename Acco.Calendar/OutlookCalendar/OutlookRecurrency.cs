using System;
//
using Microsoft.Office.Interop.Outlook;

namespace Acco.Calendar.Event
{
    class OutlookRecurrency : GenericRecurrency
    {
        public void FromOutlookRecurrencyType(OlRecurrenceType type)
        {
            switch(type)
            {
                case OlRecurrenceType.olRecursDaily:
                    Type = RecurrencyType.DAILY;
                    break;
                case OlRecurrenceType.olRecursWeekly:
                    Type = RecurrencyType.WEEKLY;
                    break;
                case OlRecurrenceType.olRecursMonthNth:
                case OlRecurrenceType.olRecursMonthly:
                    Type = RecurrencyType.MONTHLY;
                    break;
                case OlRecurrenceType.olRecursYearNth:
                case OlRecurrenceType.olRecursYearly:
                    Type = RecurrencyType.YEARLY;
                    break;
            }
        }

        public OlRecurrenceType ToOutlookRecurrencyType()
        {
            OlRecurrenceType myType = OlRecurrenceType.olRecursDaily;
            //
            switch (Type)
            {
                case RecurrencyType.DAILY:
                    myType = OlRecurrenceType.olRecursDaily;
                    break;
                case RecurrencyType.WEEKLY:
                    myType = OlRecurrenceType.olRecursWeekly;
                    break;
                case RecurrencyType.MONTHLY:
                    myType = OlRecurrenceType.olRecursMonthly;
                    break;
                case RecurrencyType.YEARLY:
                    myType = OlRecurrenceType.olRecursYearly;
                    break;
            }
            //
            return myType;
        }
    }
}
