using System.Reflection;
using System.Windows;
using System.Windows.Input;
using OpenCalendarSync.App.Tray.Properties;
using OpenCalendarSync.Lib.Manager;
using System;
using OpenCalendarSync.Lib.Database;
using System.Windows.Threading;
using Ookii.Dialogs.Wpf;

namespace OpenCalendarSync.App.Tray
{
    /// <summary>
    /// Logica di interazione per SettingsDialog.xaml
    /// </summary>
    public partial class SettingsDialog
    {
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        private string _clientId;
        private string _clientSecret;

        public SettingsDialog()
        {
            InitializeComponent();
            InternalInit();
        }

        public SettingsDialog(Hardcodet.Wpf.TaskbarNotification.TaskbarIcon trayIcon)
        {
            InitializeComponent();
            InternalInit();
            _trayIcon = trayIcon;
        }

        private void InternalInit()
        {
            _colorChanged = new bool[2];

            _clientId = Settings.Default.ClientID;
            _clientSecret = Settings.Default.ClientSecret;
            if (string.IsNullOrEmpty(_clientId))
                _clientId = GoogleToken.ClientId;
            if (string.IsNullOrEmpty(_clientSecret))
                _clientSecret = GoogleToken.ClientSecret;

            ClientIdPwdBox.Password = _clientId;
            ClientSecretPwdBox.Password = _clientSecret;

            TextColorComboBox.SelectedColorChanged += textColorComboBox_SelectedColorChanged;
            BackgroundColorComboBox.SelectedColorChanged += backgroundColorComboBox_SelectedColorChanged;

            LibraryVersionLabel.Content =   "Lib v" + Lib.Utilities.VersionHelper.LibraryVersion() +
                                            ", built " +  Lib.Utilities.VersionHelper.LibraryBuildTime().ToString("s");
            ExecutingAssemblyVersionLabel.Content = "App v" + Lib.Utilities.VersionHelper.ExecutingAssemblyVersion() +
                                                    ", built " + Lib.Utilities.VersionHelper.ExecutingAssemblyBuildTime().ToString("s");
        }

        private void btSave_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = true;
        }

        private void btCancel_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
        }

        private void slRefreshTmo_MouseWheel(object sender, MouseWheelEventArgs e)
        {
            if (e.Delta > 0) { SlRefreshTmo.Value += SlRefreshTmo.SmallChange; }
            else { SlRefreshTmo.Value -= SlRefreshTmo.SmallChange; }
        }

        private async void btReset_Click(object sender, RoutedEventArgs e)
        {
            var askForReset = MessageBox.Show("L'operazione di reset comporta:\n" +
                                                "\t1. Cancellazione database appuntamenti\n" +
                                                "\t2. Cancellazione calendario su google (opzionale)\n" +
                                                "Vuoi proseguire?",
                                                "Reset",
                                                MessageBoxButton.YesNo,
                                                MessageBoxImage.Exclamation,
                                                MessageBoxResult.No);

            if (askForReset != MessageBoxResult.Yes) return;
            if (!GoogleCalendarManager.Instance.LoggedIn)
            {
                var res = await GoogleCalendarManager.Instance.Login(_clientId, _clientSecret);
                if(!res)
                { 
                    Log.Error("Couldn't log in to google services with the provided clientId and clientSecret");
                    _trayIcon.ShowBalloonTip("Errore", "Login a servizi google non effettuato.", Hardcodet.Wpf.TaskbarNotification.BalloonIcon.Error);
                    HideBalloonAfterSeconds(6);
                    return;
                }
            }
            Reset();
        }    

        private async void Reset()
        {
            const string title = "Risultato reset";
            var text = "";
            var drop = Storage.Instance.Appointments.Drop();
            if (drop.Ok)
            {
                text += "\tDatabase appuntamenti svuotato correttamente\n";
            }
            else
            {
                text += "\tDatabase appuntamenti *NON* svuotato!\n";
                Log.Error("Failed to delete drop appointments database");
            }
            try
            {
                var res = MessageBox.Show("Vuoi cancellare il calendario su google?",
                                           "Calendario Google",
                                           MessageBoxButton.YesNo, MessageBoxImage.Exclamation, MessageBoxResult.No);
                if(res == MessageBoxResult.Yes)
                {
                    var calendarDrop = await GoogleCalendarManager.Instance.DropCurrentCalendar();
                    text += "\tCalendario su google calendar cancellato correttamente\n";
                    if(!string.IsNullOrEmpty(calendarDrop))
                        text += "\t\tDettagli operazione: " + calendarDrop + "\n";
                }
            }
            catch (Exception ex)
            {
                text += "\tCalendario su google calendar *NON* cancellato\n";
                Log.Error("Failed to delete calendar from google", ex);
            }
            Settings.Default.CalendarID = "";
            Settings.Default.Save();
            text += "\tID del calendario resettato";
            _trayIcon.ShowBalloonTip(title, text, Hardcodet.Wpf.TaskbarNotification.BalloonIcon.Info);
            HideBalloonAfterSeconds(6);
        }

        private bool[] _colorChanged;
        private readonly Hardcodet.Wpf.TaskbarNotification.TaskbarIcon _trayIcon;
        private void textColorComboBox_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<System.Windows.Media.Color> e)
        {
            if(e.NewValue != e.OldValue)
            {
                _colorChanged[0] = true;
            }
        }

        private void backgroundColorComboBox_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<System.Windows.Media.Color> e)
        {
            if (e.NewValue != e.OldValue)
            {
                _colorChanged[1] = true;
            }
        }

        private async void ChangeCalendarColor(System.Windows.Media.Color fg, System.Windows.Media.Color bg)
        {
            if(_colorChanged[0] && _colorChanged[1])
            {
                _colorChanged[0] = false;
                _colorChanged[1] = false;
                var foregroundColor =  "#" + fg.ToString().Substring(3); // remove alpha channel, we don't want that shit
                var backgroundColor = "#" + bg.ToString().Substring(3); 
                if(!GoogleCalendarManager.Instance.LoggedIn)
                {
                    var login = await GoogleCalendarManager.Instance.Login(_clientId, _clientSecret);
                }
                var calendarId = await GoogleCalendarManager.Instance.Initialize(Settings.Default.CalendarID, Settings.Default.CalendarName);
                if (Settings.Default.CalendarID != null && Settings.Default.CalendarID != calendarId)
                { 
                    Settings.Default.CalendarID = calendarId;
                }
                if(GoogleCalendarManager.Instance.LoggedIn)
                { 
                    await GoogleCalendarManager.Instance.SetCalendarColor(foregroundColor.ToLower(), backgroundColor.ToLower());
                }
                _trayIcon.ShowBalloonTip("Colore calendario", "Il colore del tuo calendario su google e' stato cambiato correttamente", Hardcodet.Wpf.TaskbarNotification.BalloonIcon.Info);
                HideBalloonAfterSeconds(6);
            }
        }

        private void HideBalloonAfterSeconds(int seconds)
        {
            var tmr = new DispatcherTimer { Interval = TimeSpan.FromSeconds(seconds) };
            tmr.Tick += delegate
            {
                tmr.Stop();
                _trayIcon.HideBalloonTip();
            };
            tmr.Start();
        }

        private void calnameTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            //todo: add summary customization here
        }

        private void textColorComboBox_Unloaded(object sender, RoutedEventArgs e)
        {
            ChangeCalendarColor(TextColorComboBox.SelectedColor, BackgroundColorComboBox.SelectedColor);
        }

        private void backgroundColorComboBox_Unloaded(object sender, RoutedEventArgs e)
        {
            ChangeCalendarColor(TextColorComboBox.SelectedColor, BackgroundColorComboBox.SelectedColor);
        }

        private void UpdatesRepositoryTextBox_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            var dialog = new VistaFolderBrowserDialog();
            if (!string.IsNullOrEmpty(Settings.Default.UpdateRepositoryPath))
                dialog.SelectedPath = Settings.Default.UpdateRepositoryPath;
            var result = dialog.ShowDialog();
            if (!result.HasValue) return;
            if (!result.Value) return;
            Settings.Default.UpdateRepositoryPath = dialog.SelectedPath;
        }
    }
}
