using System.Linq;
using MongoDB.Driver.Builders;
using OpenCalendarSync.App.Tray.Properties;
using OpenCalendarSync.Lib;
using OpenCalendarSync.Lib.Event;
using OpenCalendarSync.Lib.Manager;
using Hardcodet.Wpf.TaskbarNotification;
using log4net.Config;
using System;
using System.Reflection;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Threading;
using System.Collections.Generic;
using System.IO;
using Squirrel;
using Ookii.Dialogs.Wpf;

namespace OpenCalendarSync.App.Tray
{
    /// <summary>
    /// Logica di interazione per MainWindow.xaml
    /// </summary>
    public partial class MainWindow
    {
        private readonly DispatcherTimer _iconAnimationTimer;
        private int _currentIconIndex;
        private readonly System.Drawing.Icon[] _animationIcons;
        private readonly System.Drawing.Icon _idleIcon;
        private bool _animationStopping;

        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        static bool _showTheWelcomeWizard;

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

            _currentIconIndex = 0;
            _animationIcons    = new System.Drawing.Icon[11];
            _animationStopping = false;

            for ( var i = 0; i < 11; ++i ) {
                try
                {
                    var resName = string.Format("_{0}", i + 1);
                    _animationIcons[i] = GetAppIcon(resName, new System.Drawing.Size(32,32));
                }
                catch (Exception e)
                {
                    Log.Error("Exception", e);
                }
            }

            _idleIcon = GetAppIcon("app", new System.Drawing.Size(256, 256));
            TrayIcon.Icon = _idleIcon;

            // Create a Timer with a Normal Priority
            var timer = new DispatcherTimer {Interval = TimeSpan.FromMinutes(Settings.Default.RefreshRate)};

            // Set the callback to just show the time ticking away
            // NOTE: We are using a control so this has to run on 
            // the UI thread
            timer.Tick += delegate {
                StartSync();
            };

            timer.Start();

            _iconAnimationTimer = new DispatcherTimer {Interval = TimeSpan.FromMilliseconds(50)};

            _iconAnimationTimer.Tick += delegate
            {
                if (_animationStopping && _currentIconIndex == 0)
                {
                    _animationStopping = false;
                    _iconAnimationTimer.Stop();
                    TrayIcon.Icon = _idleIcon;
                }
                else
                {
                    _currentIconIndex = ++_currentIconIndex % 11;
                    TrayIcon.Icon = _animationIcons[_currentIconIndex];
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
            MiStatus.Header = "Sincronizzazione in corso...";

            _currentIconIndex = 0;
            _iconAnimationTimer.Start();

            await OutlookToGoogle();
            EndSync();
        }

        private void EndSync()
        {
            MiStatus.Header = "In attesa...";
            _animationStopping = true;            
        }

        private async Task OutlookToGoogle()
        {
            // settings
            var clientId = Settings.Default.ClientID;
            var secret    = Settings.Default.ClientSecret;
            var calName  = Settings.Default.CalendarName;
            var calId    = Settings.Default.CalendarID;   // note: on first startup, this is null or empty.
            if ( string.IsNullOrEmpty(clientId) )
                clientId = GoogleToken.ClientId;
            if ( string.IsNullOrEmpty(secret) )
                secret    = GoogleToken.ClientSecret;
            if (string.IsNullOrEmpty(calName))
                calName = "GVR.Meetings";
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
                    var login = await GoogleCalendarManager.Instance.Login(clientId, secret);
                }
                // initialize google calendar (i.e.: create it if it's not present, just get it if it's present)
                var googleCalId = await GoogleCalendarManager.Instance.Initialize(calId, calName);
                if (calId != googleCalId) // if the calendar ids differ
                {
                    // update the settings, so that the next time we start, we have a calendarId
                    Settings.Default.CalendarID = googleCalId;
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
            if (events != null && events.Count > 0)
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
                TrayIcon.ShowBalloonTip(title, text, BalloonIcon.Info);
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
            TrayIcon.ShowBalloonTip(title, text, BalloonIcon.Error);
        }

        private void HideBalloonAfterSeconds(int seconds)
        {
            var tmr = new DispatcherTimer {Interval = TimeSpan.FromSeconds(seconds)};
            tmr.Tick += delegate
            {
                tmr.Stop();
                //hide balloon
                TrayIcon.HideBalloonTip();                
            };
            tmr.Start();
        }

        private void miSettings_Click(object sender, RoutedEventArgs e)
        {
            var sd = new SettingsDialog(TrayIcon);
            var result = sd.ShowDialog();
            if (result.HasValue && result.Value)
            {
                Settings.Default.Save();
            }
        }

        private UpdateManager _mgr;

        private async void Window_Initialized(object sender, EventArgs e)
        {
            Log.Debug(String.Format("Window_Initialized: sender[{0}], args [{1}]", sender, e));

            // todo: avoid to login to google if there's any squirrel event
            var clientId = Settings.Default.ClientID;
            var secret = Settings.Default.ClientSecret;
            if (string.IsNullOrEmpty(clientId))
                clientId = GoogleToken.ClientId;
            if (string.IsNullOrEmpty(secret))
                secret = GoogleToken.ClientSecret;

            if(!GoogleCalendarManager.Instance.LoggedIn)
            {
                var login = await GoogleCalendarManager.Instance.Login(clientId, secret);
            }

            var repo = Settings.Default.UpdateRepositoryPath;
            if (string.IsNullOrEmpty(Settings.Default.UpdateRepositoryPath))
                repo = "null";

            using (var manager = new UpdateManager(repo, "OpenCalendarSync", FrameworkVersion.Net45))
            {
                // Note, in most of these scenarios, the app exits after this method
                // completes!
                SquirrelAwareApp.HandleEvents(
                onInitialInstall: v =>
                {
                    Log.Info(String.Format("Application installed {0}", v.ToString()));
                    TrayIcon.ShowBalloonTip("Installazione", String.Format("OpenCalendarSync {0} installato.", v.ToString()), BalloonIcon.Info);
                    HideBalloonAfterSeconds(10);
                    manager.CreateShortcutForThisExe();
                },
                onAppUpdate: v =>
                {
                    Log.Info(String.Format("Application updated to {0}, trying to upgrade settings", v.ToString()));
                    try
                    {
                        Settings.Default.Upgrade();
                        Settings.Default.Save();
                    }
                    catch (Exception ex)
                    {
                        Log.Error("Error while upgrading settings, you'll have to merge them manually", ex);
                    }
                    TrayIcon.ShowBalloonTip("Aggiornamento", String.Format("OpenCalendarSync aggiornato alla versione {0}", v.ToString()), BalloonIcon.Info);
                    manager.CreateShortcutForThisExe();
                },
                onAppUninstall: v => manager.RemoveShortcutForThisExe(),
                onFirstRun: () => _showTheWelcomeWizard = true);
            }

            if (!_showTheWelcomeWizard) return;
            using (var welcomeDialog = new TaskDialog())
            {
                welcomeDialog.Buttons.Add(new TaskDialogButton(ButtonType.Ok));
                welcomeDialog.WindowTitle = "Primo Avvio";
                welcomeDialog.MainIcon = TaskDialogIcon.Information;
                welcomeDialog.MainInstruction +=
                    "Benvenuto in OpenCalendarSync";
                welcomeDialog.Content +=
                    "Questo programma serve per importare il tuo calendario Microsoft Outlook nel servizio Google Calendar\n\n" +
                    "Per iniziare basta premere con qualsiasi tasto del mouse sull'icona che e' apparsa nella barra delle notifiche e\n" +
                    "e selezionare \"Sincronizza ora\", questo creera' un calendario su Google Calendar e procedera'\n" +
                    "con l'importazione del calendario attuale di Microsoft Outlook\n";

                welcomeDialog.ShowDialog();
            }
        }

        private ProgressDialog _updateDialog;
        private static bool _updateFinished;

        private async void MiUpdate_OnClick(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(Settings.Default.UpdateRepositoryPath))
            { 
                Log.Info("No repository path is set, won't update");
                return;
            }

            try
            {
                _mgr = new UpdateManager(Settings.Default.UpdateRepositoryPath, "OpenCalendarSync", FrameworkVersion.Net45);
                var updateInfo = await _mgr.CheckForUpdate();
                if (!updateInfo.ReleasesToApply.Any())
                {
                    Log.Info("No newer version found");
                    return;
                }

                Log.Info(String.Format("Newer version found => [{0}]", updateInfo.FutureReleaseEntry.EntryAsString));

                var updateAvailableDialog = new TaskDialog
                {
                    MainIcon = TaskDialogIcon.Information,
                    WindowTitle = "Nuova versione disponibile",
                    MainInstruction = "Procedere con l'installazione?",
                    Content = String.Format("Versione {0} disponibile.\nSelezionare una delle due opzioni.", updateInfo.FutureReleaseEntry.EntryAsString)
                };
                updateAvailableDialog.Buttons.Add(new TaskDialogButton(ButtonType.Yes));
                updateAvailableDialog.Buttons.Add(new TaskDialogButton(ButtonType.No));
                var userDecision = updateAvailableDialog.Show();

                if (userDecision.ButtonType != ButtonType.Yes)
                { 
                    Log.Info("User has selected to abort the update");
                    return;
                }

                Log.Info("User has selected to proceed with the update");

                _updateDialog = new ProgressDialog();
                _updateDialog.WindowTitle += "Aggiornamento a nuova versione";
                _updateDialog.Text += "Installazione in corso dell'ultima versione...";
                _updateDialog.DoWork += (o, args) =>
                {
                    _updateFinished = false;
                    var updateAppTask = _mgr.UpdateApp(i =>
                    {
                        if ((i >= 0 && i <= 100) && !_updateFinished)
                        { 
                            _updateDialog.ReportProgress(i);
                        }
                    });
                    Log.Debug(String.Format("{0} applied.", updateAppTask.Result.EntryAsString));
                };
                _updateDialog.RunWorkerCompleted += (o, args) =>
                {
                    Log.Info("Finished updating app to the latest version.");
                    _updateFinished = true;
                    _mgr = null;
                };
                _updateDialog.Show();
            }
            catch (Exception ex)
            {
                Log.Error("Failed to check or apply update", ex);
                TrayIcon.ShowBalloonTip("Errore", "Ricerca aggiornamenti o applicazione aggiornamento fallita", BalloonIcon.Error);
                HideBalloonAfterSeconds(10);
            }
        }
    }
}
