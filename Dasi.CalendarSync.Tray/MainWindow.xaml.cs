using System.Linq;
using Acco.Calendar;
using Acco.Calendar.Manager;
using Acco.Calendar.Database;
using Google;
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
using MongoDB;

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

            var my_asm = Assembly.GetExecutingAssembly();

            for ( int i = 0; i < 11; ++i ) {
                try
                {
                    var res_name = string.Format("_{0}", i + 1);
                    var obj = Properties.Resources.ResourceManager.GetObject(res_name, Properties.Resources.Culture);
                    animation_icons[i] = (System.Drawing.Icon)(obj);
                }
                catch (Exception e)
                {
                    Log.Error("Exception", e);
                }
            }

            idle_icon = Properties.Resources.calendar;

            // Create a Timer with a Normal Priority
            _timer = new DispatcherTimer();
            _timer.Interval = TimeSpan.FromHours(1);
            //_timer.Interval = TimeSpan.FromSeconds(10);

            // Set the callback to just show the time ticking away
            // NOTE: We are using a control so this has to run on 
            // the UI thread
            _timer.Tick += delegate {
                StartSync();
            };

            _timer.Start();

            _icon_animation_timer = new DispatcherTimer();
            _icon_animation_timer.Interval = TimeSpan.FromMilliseconds(100);

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
            var client_id = Settings.Default.ClientID;
            var secret    = Settings.Default.ClientSecret;
            var cal_name  = Settings.Default.CalendarName;

            if ( string.IsNullOrEmpty(client_id) )
                client_id = GoogleToken.ClientId;
            if ( string.IsNullOrEmpty(secret) )
                secret    = GoogleToken.ClientSecret;
            if (string.IsNullOrEmpty(cal_name))
                cal_name = "GVR.Meetings";

            try
            {
                // take events from outlook and push em to google
                var calendar = await OutlookCalendarManager.Instance.PullAsync() as GenericCalendar;
                //
                var isLoggedIn = await GoogleCalendarManager.Instance.Initialize(client_id, secret, cal_name);
                if (isLoggedIn) //logged in to google, go on!
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
                    catch(Exception ex)
                    {
                        SyncFailure(ex.Message);
                    }
                }
            }
            catch (GoogleApiException ex)
            {
                Log.Error("GoogleApiException", ex);
                SyncFailure(ex.Message);
            }
            catch(Exception ex)
            {
                Log.Error("Exception", ex);
                SyncFailure(ex.Message);
            }
        }

        private void SyncSuccess(IEnumerable<UpdateOutcome> pushedEvents)
        {
            string title = "Risultato sincronizzazione";
            string text  = "La sincronizzazione e' terminata con successo";
            var events = pushedEvents as List<UpdateOutcome>;
            if (events != null && (events.Count(e => e.Successful) > 0))
            { 
                text += "\n" + String.Format("{0} eventi aggiunti al calendario", events.Count(e => e.Successful));
                //show balloon with built-in icon
                trayIcon.ShowBalloonTip(title, text, BalloonIcon.Info);
                // hide baloon in 3 seconds
                HideBalloonAfterSeconds(3);
            }
        }

        private void SyncFailure(string message)
        {
            string title = "Risultato sincronizzazione";
            string text = "La sincronizzazione e' fallita\n";
            string details = message;
            text += details;

            //show balloon with built-in icon
            trayIcon.ShowBalloonTip(title, text, BalloonIcon.Error);
        }

        private void HideBalloonAfterSeconds(int seconds)
        {
            var tmr = new DispatcherTimer();
            tmr.Interval = TimeSpan.FromSeconds(seconds);
            tmr.Tick += delegate
            {
                tmr.Stop();
                //hide balloon
                trayIcon.HideBalloonTip();                
            };
            tmr.Start();
        }

        private async void miReset_Click(object sender, RoutedEventArgs e)
        {
            const string title = "Risultato reset";
            var text = "";
            var drop = Storage.Instance.Appointments.Drop();
            if(drop.Ok)            
            {
                text += "Database appuntamenti svuotato correttamente\n";
            }
            else
            {
                text += "Database appuntamenti *NON* svuotato!\n";
                Log.Error("Failed to delete drop appointments database");
            }
            try
            {
                var clientId = GoogleToken.ClientId;
                var secret = GoogleToken.ClientSecret;
                var calName = Settings.Default.CalendarName;
                //

                var isLoggedIn = await GoogleCalendarManager.Instance.Initialize(clientId, secret, calName);
                if (isLoggedIn) //logged in to google, go on!
                {
                    var calendarDrop = await GoogleCalendarManager.Instance.DropCurrentCalendar();
                    text += "Calendario su google calendar cancellato correttamente\n";
                    text += " Dettagli operazione: " + calendarDrop + "\n";
                }
            }
            catch(Exception ex)
            {
                text += "Calendario su google calendar *NON* cancellato\n";
                Log.Error("Failed to delete calendar from google", ex);
            }
            //
            try
            {
                
                File.Delete("googlecalendar.settings");
                text += "Impostazioni di google calendar cancellate correttamente";
            }
            catch(Exception ex)
            {
                text += "Impostazioni di google calendar *NON* cancellate!";
                Log.Error("Failed to delete googlecalendar.settings", ex);
            }
            //
            trayIcon.ShowBalloonTip(title, text, BalloonIcon.Warning);
        }

        private void miSettings_Click(object sender, RoutedEventArgs e)
        {
            var sd = new SettingsDialog();
            bool? result = sd.ShowDialog();
            if (result.HasValue && result.Value)
            {
                Properties.Settings.Default.Save();
                // save settings! YEAH
                // write your code here!
            }
        }
    }
}
