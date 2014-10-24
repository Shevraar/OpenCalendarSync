using System.Linq;
using Acco.Calendar;
using Acco.Calendar.Event;
using Acco.Calendar.Manager;
using Acco.Calendar.Database;
using Hardcodet.Wpf.TaskbarNotification;
using log4net.Config;
using System;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using System.Collections.Generic;
using Dasi.CalendarSync.Tray.Properties;
using System.IO;
using Squirrel;
using Ookii.Dialogs.Wpf;

namespace Dasi.CalendarSync.Tray
{
    /// <summary>
    /// Logica di interazione per MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private DispatcherTimer _timer;
        private DispatcherTimer _icon_animation_timer;
        private int current_icon_index;
        private System.Drawing.Icon[] animation_icons;
        private System.Drawing.Icon idle_icon;
        private bool animation_stopping;

        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        static bool ShowTheWelcomeWizard;

        public MainWindow()
        {
            XmlConfigurator.Configure(); //only once
            Log.Info("Application is starting");

            try
            {
                InitializeComponent();
            }
            catch (Exception ex)
            {
                Log.Error("Exception", ex);    
            }

            current_icon_index = 0;
            animation_icons    = new System.Drawing.Icon[11];
            animation_stopping = false;

            for ( var i = 0; i < 11; ++i ) {
                try
                {
                    var res_name = string.Format("_{0}", i + 1);
                    animation_icons[i] = GetAppIcon(res_name, new System.Drawing.Size(32,32));
                }
                catch (Exception e)
                {
                    Log.Error("Exception", e);
                }
            }

            idle_icon = GetAppIcon("app", new System.Drawing.Size(256, 256));
            trayIcon.Icon = idle_icon;

            // Create a Timer with a Normal Priority
            _timer = new DispatcherTimer {Interval = TimeSpan.FromMinutes(Settings.Default.RefreshRate)};

            // Set the callback to just show the time ticking away
            // NOTE: We are using a control so this has to run on 
            // the UI thread
            _timer.Tick += delegate {
                StartSync();
            };

            _timer.Start();

            _icon_animation_timer = new DispatcherTimer {Interval = TimeSpan.FromMilliseconds(50)};

            _icon_animation_timer.Tick += delegate
            {
                if (animation_stopping && current_icon_index == 0)
                {
                    animation_stopping = false;
                    _icon_animation_timer.Stop();
                    trayIcon.Icon = idle_icon;
                }
                else
                {
                    current_icon_index = ++current_icon_index % 11;
                    trayIcon.Icon = animation_icons[current_icon_index];
                }
            };
        }

        private System.Drawing.Icon GetAppIcon(string name, System.Drawing.Size sz)
        {
            // se esiste la icona "doge" usa quella :)
            var dogeFile = Path.Combine("doge", name + ".ico");
            if (File.Exists(dogeFile))
                return new System.Drawing.Icon(dogeFile, sz);
            return (System.Drawing.Icon)Properties.Resources.ResourceManager.GetObject(name, Properties.Resources.Culture);
        }

        private void miExit_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

        private void miSync_Click(object sender, RoutedEventArgs e)
        {
            StartSync();
        }

        private async void StartSync()
        {
            miStatus.Header = "Sincronizzazione in corso...";

            current_icon_index = 0;
            _icon_animation_timer.Start();

            await OutlookToGoogle();
            EndSync();
        }

        private void EndSync()
        {
            miStatus.Header = "In attesa...";
            animation_stopping = true;            
        }

        private async Task OutlookToGoogle()
        {
            // settings
            var client_id = Settings.Default.ClientID;
            var secret    = Settings.Default.ClientSecret;
            var cal_name  = Settings.Default.CalendarName;
            var cal_id    = Settings.Default.CalendarID;   // note: on first startup, this is null or empty.
            if ( string.IsNullOrEmpty(client_id) )
                client_id = GoogleToken.ClientId;
            if ( string.IsNullOrEmpty(secret) )
                secret    = GoogleToken.ClientSecret;
            if (string.IsNullOrEmpty(cal_name))
                cal_name = "GVR.Meetings";
            //
            ICalendar calendar;
            //
            try
            {
                // take events from outlook and push em to google
                calendar = await OutlookCalendarManager.Instance.PullAsync() as GenericCalendar;
            }
            catch (Exception ex)
            {
                SyncFailure(ex.Message);
                return;
            }
            //
            try
            {
                // if I'm not logged in
                if (!GoogleCalendarManager.Instance.LoggedIn)
                {
                    var login = await GoogleCalendarManager.Instance.Login(client_id, secret);
                }
                // initialize google calendar (i.e.: create it if it's not present, just get it if it's present)
                var google_cal_id = await GoogleCalendarManager.Instance.Initialize(cal_id, cal_name);
                if (cal_id != google_cal_id) // if the calendar ids differ
                {
                    // update the settings, so that the next time we start, we have a calendarId
                    Settings.Default.CalendarID = google_cal_id;
                    Settings.Default.Save();
                }
                //logged in to google, go on!
                if (GoogleCalendarManager.Instance.LoggedIn) 
                {
                    try
                    {
                        var ret = await GoogleCalendarManager.Instance.PushAsync(calendar);
                        SyncSuccess(ret);
                    }
                    catch (PushException ex)
                    {
                        SyncFailure(ex.Message);
                    }
                    catch (Exception ex)
                    {
                        SyncFailure(ex.Message);
                    }
                }
            }
            catch(Exception ex)
            {
                Log.Error("Exception", ex);
                SyncFailure(ex.Message);
            }
        }

        private void SyncSuccess(IEnumerable<UpdateOutcome> pushedEvents)
        {
            const string title = "Risultato sincronizzazione";
            var text  = "La sincronizzazione e' terminata con successo";
            var events = pushedEvents as List<UpdateOutcome>;
            if (events.Count > 0)
            {
                if(events.Count(e => e.Successful && e.Event.Action == EventAction.Add) > 0)
                {
                    text += "\n" + String.Format("{0} eventi aggiunti al calendario", events.Count(e => e.Successful && e.Event.Action == EventAction.Add));
                }
                if (events.Count(e => e.Successful && e.Event.Action == EventAction.Update) > 0)
                {
                    text += "\n" + String.Format("{0} eventi aggiornati", events.Count(e => e.Successful && e.Event.Action == EventAction.Update));
                }
                if (events.Count(e => e.Event.Action == EventAction.Remove) > 0)
                {
                    text += "\n" + String.Format("{0} eventi rimossi", events.Count(e => e.Event.Action == EventAction.Remove));
                }
                trayIcon.ShowBalloonTip(title, text, BalloonIcon.Info);
                HideBalloonAfterSeconds(10);
            }
        }

        private void SyncFailure(string message)
        {
            const string title = "Risultato sincronizzazione";
            var text = "La sincronizzazione e' fallita\n";
            var details = message;
            text += details;

            //show balloon with built-in icon
            trayIcon.ShowBalloonTip(title, text, BalloonIcon.Error);
        }

        private void HideBalloonAfterSeconds(int seconds)
        {
            var tmr = new DispatcherTimer {Interval = TimeSpan.FromSeconds(seconds)};
            tmr.Tick += delegate
            {
                tmr.Stop();
                //hide balloon
                trayIcon.HideBalloonTip();                
            };
            tmr.Start();
        }

        private void miSettings_Click(object sender, RoutedEventArgs e)
        {
            var sd = new SettingsDialog(trayIcon);
            var result = sd.ShowDialog();
            if (result.HasValue && result.Value)
            {
                Settings.Default.Save();
            }
        }

        private async void Window_Initialized(object sender, EventArgs e)
        {
            if (string.IsNullOrEmpty(Settings.Default.UpdateRepositoryPath)) return;
            using (var mgr = new UpdateManager(Settings.Default.UpdateRepositoryPath, "OpenCalendarSync", FrameworkVersion.Net45))
            {
                // Note, in most of these scenarios, the app exits after this method
                // completes!
                SquirrelAwareApp.HandleEvents(
                  onInitialInstall: v => mgr.CreateShortcutForThisExe(),
                  onAppUpdate: v => mgr.CreateShortcutForThisExe(),
                  onAppUninstall: v => mgr.RemoveShortcutForThisExe(),
                  onFirstRun: () => ShowTheWelcomeWizard = true);
            }
            //
            var client_id = Settings.Default.ClientID;
            var secret = Settings.Default.ClientSecret;
            if (string.IsNullOrEmpty(client_id))
                client_id = GoogleToken.ClientId;
            if (string.IsNullOrEmpty(secret))
                secret = GoogleToken.ClientSecret;

            if(!GoogleCalendarManager.Instance.LoggedIn)
            {
                var login = await GoogleCalendarManager.Instance.Login(client_id, secret);
            }
        }

        private void trayIcon_Loaded(object sender, RoutedEventArgs e)
        {
            if (!ShowTheWelcomeWizard) return;
            using(var welcomeDialog = new TaskDialog())
            {
                welcomeDialog.Buttons.Add(new TaskDialogButton(ButtonType.Ok));
                welcomeDialog.WindowTitle = "Primo avvio di OpenCalendarSync (o aggiornamento)";
                welcomeDialog.Content += "Questo programma serve per importare il tuo calendario Microsoft Outlook nel servizio Google Calendar\n";
                welcomeDialog.Content += "Successivamente sara' possibile il colore da assegnare al calendario, cosi' come il nome\n";
                var res = welcomeDialog.ShowDialog();
                if (res.ButtonType != ButtonType.Ok) return;
            }
        }
    }
}
