using Acco.Calendar;
using Acco.Calendar.Manager;
using Google;
using Hardcodet.Wpf.TaskbarNotification;
using log4net.Config;
using System;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using System.Collections.Generic;

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
            InitializeComponent();

            XmlConfigurator.Configure(); //only once
            Log.Info("Application is starting");

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
            var client_id = GoogleToken.ClientId;
            var secret    = GoogleToken.ClientSecret;
            var cal_name  = "GVR.Meetings";

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
                        //todo: do something with the list returned from PushAsync.
                    }
                    catch (PushException ex)
                    {
                        SyncFailure();
                    }
                }
            }
            catch (GoogleApiException ex)
            {
                SyncFailure();
            }
        }

        private void SyncSuccess(IEnumerable<PushedEvent> pushedEvents)
        {
            string title = "Risultato sincronizzazione";
            string text  = "La sincronizzazione e' terminata con successo";
            var events = pushedEvents as List<PushedEvent>;
            if (events != null)
                text += "\n" + String.Format("{0} eventi aggiunti al calendario", events.Count);

            //show balloon with built-in icon
            trayIcon.ShowBalloonTip(title, text, BalloonIcon.Info);
            
            // hide baloon in 3 seconds
            HideBalloonAfterSeconds(3);
        }

        private void SyncFailure()
        {
            string title = "Risultato sincronizzazione";
            string text = "La sincronizzazione e' fallita";

            //show balloon with built-in icon
            trayIcon.ShowBalloonTip(title, text, BalloonIcon.Error);

            // hide baloon in 3 seconds
            HideBalloonAfterSeconds(3);
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
    }
}
