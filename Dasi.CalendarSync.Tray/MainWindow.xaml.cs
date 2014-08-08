using Acco.Calendar;
using Acco.Calendar.Manager;
using Google;
using Hardcodet.Wpf.TaskbarNotification;
using log4net.Config;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Resources;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;

namespace Dasi.CalendarSync.Tray
{
    /// <summary>
    /// Logica di interazione per MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
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
            _timer.Tick += new EventHandler(delegate(object s, EventArgs a)
            {
                StartSync();
            });

            _timer.Start();

            _icon_animation_timer = new DispatcherTimer();
            _icon_animation_timer.Interval = TimeSpan.FromMilliseconds(100);

            _icon_animation_timer.Tick += new EventHandler(delegate(object s, EventArgs a)
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
            });

            
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
            var client_id = "498986356901-ff0h9agji4iqkfniec38cga49k310p92.apps.googleusercontent.com";
            var secret    = "a3d1FPuuQ738UoxkMzk7T8LN";
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
                        SyncSuccess();
                        /*if (ret)
                        {
                            SyncSuccess();
                        }
                        else { SyncFailure(); }*/
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

        private void SyncSuccess()
        {
            string title = "Risultato sincronizzazione";
            string text  = "La sincronizzazione e' terminata con successo";

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
            tmr.Tick += new EventHandler(delegate(object s, EventArgs a)
            {
                tmr.Stop();
                //hide balloon
                trayIcon.HideBalloonTip();                
            });
            tmr.Start();
        }
    }
}
