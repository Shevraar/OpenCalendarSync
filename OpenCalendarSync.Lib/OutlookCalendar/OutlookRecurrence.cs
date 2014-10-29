//

using System.Diagnostics.Eventing.Reader;
using System.Linq;
using System.Runtime.InteropServices;
using RecPatt = DDay.iCal.RecurrencePattern;
using iCal = DDay.iCal;

//
using Microsoft.Office.Interop.Outlook;
using System;
using System.Collections.Generic;
using OutlookRecurrencePattern = Microsoft.Office.Interop.Outlook.RecurrencePattern;
using RecurrenceException = Microsoft.Office.Interop.Outlook.Exception;

namespace OpenCalendarSync.Lib.Event
{
    public class OutlookRecurrence : GenericRecurrence
    {
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);
        public override void Parse<T>(T rules)
        {
            if (rules is OutlookRecurrencePattern)
            {
                Log.Info("Parsing OutlookRecurrencePattern...");
                var outlookRp = rules as OutlookRecurrencePattern;
                Log.Info(String.Format("Recurrence type is [{0}]", outlookRp.RecurrenceType));
                switch (outlookRp.RecurrenceType)
                {
                    case OlRecurrenceType.olRecursDaily:
                        {
                            RecurrencePattern = new RecPatt(iCal.FrequencyType.Daily, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = iCal.RecurrenceRestrictionType.Default
                            };
                            if (outlookRp.Occurrences > 0) { RecurrencePattern.Count = outlookRp.Occurrences; }
                            else if (outlookRp.PatternEndDate > DateTime.Now) { RecurrencePattern.Until = outlookRp.PatternEndDate; }
                        }
                        break;

                    case OlRecurrenceType.olRecursWeekly:
                        {
                            RecurrencePattern = new RecPatt(iCal.FrequencyType.Weekly, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = iCal.RecurrenceRestrictionType.NoRestriction
                            };
                            if (outlookRp.Occurrences > 0) { RecurrencePattern.Count = outlookRp.Occurrences; }
                            else if (outlookRp.PatternEndDate > DateTime.Now) { RecurrencePattern.Until = outlookRp.PatternEndDate; }
                            RecurrencePattern.ByDay = ExtractDaysOfWeek(outlookRp.DayOfWeekMask);
                        }
                        break;

                    case OlRecurrenceType.olRecursMonthly:
                        {
                            // warning: untested
                            // see also: http://msdn.microsoft.com/en-us/library/office/aa211012(v=office.11).aspx
                            RecurrencePattern = new RecPatt(iCal.FrequencyType.Monthly, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = iCal.RecurrenceRestrictionType.NoRestriction
                            };
                            // how many times this is going to occur
                            if (outlookRp.Occurrences > 0) { RecurrencePattern.Count = outlookRp.Occurrences; }
                            else if (outlookRp.DayOfMonth > 0) { RecurrencePattern.ByMonthDay = new List<int> { outlookRp.DayOfMonth }; }
                        }
                        break;

                    case OlRecurrenceType.olRecursMonthNth:
                        {
                            // see also: http://msdn.microsoft.com/en-us/library/office/aa211012(v=office.11).aspx
                            RecurrencePattern = new iCal.RecurrencePattern(iCal.FrequencyType.Monthly, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = iCal.RecurrenceRestrictionType.NoRestriction
                            };
                            // how many times this is going to occur
                            if (outlookRp.Occurrences > 0) { RecurrencePattern.Count = outlookRp.Occurrences; }
                            // instance states e.g.: "The Nth Tuesday"
                            if (outlookRp.Instance > 0) { RecurrencePattern.BySetPosition = new List<int> { outlookRp.Instance }; }
                            if (outlookRp.DayOfWeekMask > 0) { RecurrencePattern.ByDay = ExtractDaysOfWeek(outlookRp.DayOfWeekMask); }
                        }
                        break;

                    case OlRecurrenceType.olRecursYearly:
                        {
                            // warning: untested
                            RecurrencePattern = new RecPatt(iCal.FrequencyType.Yearly, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = iCal.RecurrenceRestrictionType.NoRestriction
                            };
                            if (outlookRp.Occurrences > 0) { RecurrencePattern.Count = outlookRp.Occurrences; }
                            else if (outlookRp.PatternEndDate > DateTime.Now) { RecurrencePattern.Until = outlookRp.PatternEndDate; }
                            if (outlookRp.DayOfWeekMask > 0) { RecurrencePattern.ByDay = ExtractDaysOfWeek(outlookRp.DayOfWeekMask); }
                            else if (outlookRp.DayOfMonth > 0) { RecurrencePattern.ByMonthDay = new List<int> { outlookRp.DayOfMonth }; }
                        }
                        break;

                    case OlRecurrenceType.olRecursYearNth:
                        {
                            // warning: untested
                            RecurrencePattern = new RecPatt(iCal.FrequencyType.Yearly, outlookRp.Interval)
                            {
                                FirstDayOfWeek = DayOfWeek.Monday,
                                RestrictionType = iCal.RecurrenceRestrictionType.NoRestriction
                            };
                            if (outlookRp.Occurrences > 0) { RecurrencePattern.Count = outlookRp.Occurrences; }
                            else if (outlookRp.PatternEndDate > DateTime.Now) { RecurrencePattern.Until = outlookRp.PatternEndDate; }
                            if (outlookRp.DayOfWeekMask > 0) { RecurrencePattern.ByDay = ExtractDaysOfWeek(outlookRp.DayOfWeekMask); }
                            else if (outlookRp.DayOfMonth > 0) { RecurrencePattern.ByMonthDay = new List<int> { outlookRp.DayOfMonth }; }
                        }
                        break;
                }
                // EXDATE and RDATE down here
                if(outlookRp.Exceptions.Count > 0) // there are some exceptions to the occurrences found above.
                {
                    //todo: add exception parsing here... then save it somewhere.
                    Exdate += "EXDATE:";
                    var excludedDatesList = new List<string>();
                    var includedDatesList = new List<string>();
                    foreach(RecurrenceException recurrenceException in outlookRp.Exceptions)
                    {
                        // old date
                        var oldDate = recurrenceException.OriginalDate;
                        // new date
                        DateTime? newStartDate = null;
                        DateTime? newEndDate = null;
                        try
                        {
                            newStartDate = recurrenceException.AppointmentItem.Start;
                            newEndDate = recurrenceException.AppointmentItem.End;
                        }
                        catch (COMException ex)
                        {
                            Log.Debug("Couldn't retrieve any new date for the deleted event", ex);
                        }
                        
                        //exdate     = "EXDATE" exdtparam ":" exdtval *("," exdtval) CRLF
                        //EXDATE:19960402T010000Z,19960403T010000Z,19960404T010000Z
                        excludedDatesList.Add(String.Format("{0}T{1}Z", oldDate.ToUniversalTime().ToString("yyyyMMdd"),
                                                                        oldDate.ToUniversalTime().ToString("hhmmss")));
                        if (newStartDate.HasValue && newEndDate.HasValue) // we got a PERIOD value type - http://www.kanzaki.com/docs/ical/period.html
                        {
                            includedDatesList.Add(String.Format("{0}T{1}Z/{2}T{3}Z",newStartDate.Value.ToUniversalTime().ToString("yyyyMMdd"),
                                                                                    newStartDate.Value.ToUniversalTime().ToString("hhmmss"),
                                                                                    newEndDate.Value.ToUniversalTime().ToString("yyyyMMdd"),
                                                                                    newEndDate.Value.ToUniversalTime().ToString("hhmmss")));    
                        }
                        else if(newStartDate.HasValue)
                        {
                            includedDatesList.Add(String.Format("{0}T{1}Z", newStartDate.Value.ToString("yyyyMMdd"),
                                                                            newStartDate.Value.ToString("hhmmss")));
                        }
                    }
                    Exdate += String.Join(",", excludedDatesList.ToArray());
                }
                Log.Debug(String.Format("Recurrence pattern is [{0}]", string.Join(";", Pattern.ToArray())));
            }
            else
            {
                throw new RecurrenceParseException("OutlookRecurrence: Unsupported type.", typeof(T));
            }
        }

        private static List<iCal.IWeekDay> ExtractDaysOfWeek(OlDaysOfWeek outlookDaysOfWeekMask)
        {
            Log.Info(String.Format("Extracting days of week from outlookDaysOfWeekMask [{0}]", outlookDaysOfWeekMask));
            var myDaysOfWeekMask = new List<iCal.IWeekDay>();
            //
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olMonday))
            {
                Log.Debug("Monday");
                myDaysOfWeekMask.Add(new iCal.WeekDay
                {
                    DayOfWeek = DayOfWeek.Monday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olTuesday))
            {
                Log.Debug("Tuesday");
                myDaysOfWeekMask.Add(new iCal.WeekDay
                {
                    DayOfWeek = DayOfWeek.Tuesday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olWednesday))
            {
                Log.Debug("Wednesday");
                myDaysOfWeekMask.Add(new iCal.WeekDay
                {
                    DayOfWeek = DayOfWeek.Wednesday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olThursday))
            {
                Log.Debug("Thursday");
                myDaysOfWeekMask.Add(new iCal.WeekDay
                {
                    DayOfWeek = DayOfWeek.Thursday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olFriday))
            {
                Log.Debug("Friday");
                myDaysOfWeekMask.Add(new iCal.WeekDay
                {
                    DayOfWeek = DayOfWeek.Friday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olSaturday))
            {
                Log.Debug("Saturday");
                myDaysOfWeekMask.Add(new iCal.WeekDay
                {
                    DayOfWeek = DayOfWeek.Saturday
                });
            }
            if (outlookDaysOfWeekMask.HasFlag(OlDaysOfWeek.olSunday))
            {
                Log.Debug("Sunday");
                myDaysOfWeekMask.Add(new iCal.WeekDay
                {
                    DayOfWeek = DayOfWeek.Sunday
                });
            }
            return myDaysOfWeekMask;
        }
    }
}