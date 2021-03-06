OpenCalendarSync
============

A tool to help you manage calendar imports from Microsoft Outlook to Google Calendar.

___

System Requirements
===================

| Requirement | Download |
|-------------|----------|
| .Net Framework 4.5 | http://www.microsoft.com/en-US/download/details.aspx?id=30653 |
| MongoDB | http://www.mongodb.org/downloads |

### Notes

You will need to setup MongoDB accordingly to http://docs.mongodb.org/manual/tutorial/install-mongodb-on-windows/ or to http://docs.mongodb.org/manual/tutorial/install-mongodb-on-windows/#manually-create-windows-service otherwise the library won't work.

Also, to compile the library you'll also have to install http://msdn.microsoft.com/en-us/library/15s06t57.aspx (Microsoft Office Interop Assemblies) which are used by the OutlookCalendarManager to interact with Outlook and its calendar service.

___

How the code works
============
Each calendar "manager" derives from ``ICalendarManager`` interface, that forces you to implement these generic functions:

* ``Push``
* ``Pull`` and ``Pull`` with begin date and end date.
* and their **Async** versions (``PushAsync`` and ``PullAsync``).

### Simple usage
Each calendar manager provides an instance (singleton) that lets you pull its calendar into a GenericCalendar object, which is then pushable inside a new calendar manager, without any further modifications:

```C#
// take events from outlook and push em to google
var calendar = await OutlookCalendarManager.Instance.PullAsync() as GenericCalendar;
var googleCalId = await GoogleCalendarManager.Instance.Initialize(calId, calName);
if(!GoogleCalendarManager.Instance.LoggedIn)
{
	var login = await GoogleCalendarManager.Instance.Login(clientId, secret);
}
if (login) //logged in to google, go on!
{
	var ret = await GoogleCalendarManager.Instance.PushAsync(calendar);
	if (ret) Log.Info("Success");
}
```
### Subscription

If you want to integrate your application inside a service or whatsoever a timed event is provided by using a subscription policy:

```C#
var googleCalId = await GoogleCalendarManager.Instance.Initialize(calId, calName);
if(!GoogleCalendarManager.Instance.LoggedIn)
{
	var login = await GoogleCalendarManager.Instance.Login(clientId, secret);
}
OutlookCalendarManager.Instance.Subscribers = new List<ICalendarManager>
{
	GoogleCalendarManager.Instance
};
if (login) { OutlookCalendarManager.Instance.StartLookingForChanges(TimeSpan.FromSeconds(10)); }
```

``OutlookCalendarManager`` now holds a list of subscribed calendar managers and every **N** milliseconds/seconds/minutes/hours/whatever (see arguments for ``StartLookingForChanges``)  it will look for changes inside their managed calendar and push the changes inside their subsribed calendar managers.

Status
======

* **Outlook** => **Google** calendar import works successfully.
* **Google** => **Outlook** calendar import is not yet implemented.

Contribute
==========

If you have any calendar service you would like to implement feel free to contribute.