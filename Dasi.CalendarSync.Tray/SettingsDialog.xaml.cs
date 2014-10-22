using System.Linq;
using System.Reflection;
using System.Windows;
using System.Windows.Input;
using Acco.Calendar.Manager;
using Dasi.CalendarSync.Tray.Properties;
using System;
using Acco.Calendar.Database;
using System.Windows.Threading;
using MongoDB.Driver;
using Squirrel;
using OldForms = System.Windows.Forms;

namespace Dasi.CalendarSync.Tray
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

            VersionLabel.Content = "v" + Acco.Calendar.Utilities.VersionHelper.GetCurrentVersion();
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
                var initialize = await GoogleCalendarManager.Instance.Initialize(Settings.Default.CalendarID, Settings.Default.CalendarName);
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
                //hide balloon
                _trayIcon.HideBalloonTip();
            };
            tmr.Start();
        }

        private void textColorComboBox_LostFocus(object sender, RoutedEventArgs e)
        {
            //changeCalendarColor(textColorComboBox.SelectedColor, backgroundColorComboBox.SelectedColor);
        }

        private void backgroundColorComboBox_LostFocus(object sender, RoutedEventArgs e)
        {
            //changeCalendarColor(textColorComboBox.SelectedColor, backgroundColorComboBox.SelectedColor);
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

        private async void Window_Loaded(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(Settings.Default.UpdateRepositoryPath)) return;

            using (var mgr = new UpdateManager(Settings.Default.UpdateRepositoryPath, "OpenCalendarSync", FrameworkVersion.Net45))
            {
                var updateInfo = await mgr.CheckForUpdate();
                if (!updateInfo.ReleasesToApply.Any()) return;

                var ret = MessageBox.Show("Nuova versione disponibile", "Nuova Versione", MessageBoxButton.YesNo,
                    MessageBoxImage.Information, MessageBoxResult.Yes);

                if (ret != MessageBoxResult.Yes) return;
                var x = await mgr.UpdateApp();
                Log.Debug(x);
            }
        }

        private void UpdatesRepositoryTextBox_PreviewMouseUp(object sender, MouseButtonEventArgs e)
        {
            var dialog = new OldForms.FolderBrowserDialog();
            var result = dialog.ShowDialog();
            if (result != OldForms.DialogResult.OK) return;
            Settings.Default.UpdateRepositoryPath = dialog.SelectedPath;
            Settings.Default.Save();
        }
    }
}
