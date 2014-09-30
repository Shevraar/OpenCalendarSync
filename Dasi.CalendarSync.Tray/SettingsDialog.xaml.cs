using System.Reflection;
using System.Windows;
using System.Windows.Input;
using Acco.Calendar.Manager;
using Dasi.CalendarSync.Tray.Properties;
using System;
using System.Threading.Tasks;
using Acco.Calendar.Database;
using System.IO;
using System.Windows.Threading;

namespace Dasi.CalendarSync.Tray
{
    /// <summary>
    /// Logica di interazione per SettingsDialog.xaml
    /// </summary>
    public partial class SettingsDialog
    {
        private static readonly log4net.ILog Log = log4net.LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);
        private string clientId;
        private string clientSecret;

        public SettingsDialog()
        {
            InitializeComponent();
            InternalInit();
        }

        public SettingsDialog(Hardcodet.Wpf.TaskbarNotification.TaskbarIcon trayIcon)
        {
            InitializeComponent();
            InternalInit();
            this.trayIcon = trayIcon;
        }

        private void InternalInit()
        {
            colorChanged = new bool[2];

            clientId = Settings.Default.ClientID;
            clientSecret = Settings.Default.ClientSecret;
            if (string.IsNullOrEmpty(clientId))
                clientId = GoogleToken.ClientId;
            if (string.IsNullOrEmpty(clientSecret))
                clientSecret = GoogleToken.ClientSecret;

            textColorComboBox.SelectedColorChanged += textColorComboBox_SelectedColorChanged;
            backgroundColorComboBox.SelectedColorChanged += backgroundColorComboBox_SelectedColorChanged;
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
            if (e.Delta > 0) { slRefreshTmo.Value += slRefreshTmo.SmallChange; }
            else { slRefreshTmo.Value -= slRefreshTmo.SmallChange; }
        }

        private async void btReset_Click(object sender, RoutedEventArgs e)
        {
            var askForReset = MessageBox.Show("L'operazione di reset comporta:\n" +
                                                "\t1. Cancellazione database appuntamenti\n" +
                                                "\t2. Cancellazione calendario su google (opzionale)\n" +
                                                "\t3. Cancellazione google.settings\n" +
                                                "Vuoi proseguire?",
                                                "Reset",
                                                MessageBoxButton.YesNo,
                                                MessageBoxImage.Exclamation,
                                                MessageBoxResult.No);

            if (askForReset == MessageBoxResult.Yes)
            {
                if (!GoogleCalendarManager.Instance.LoggedIn)
                {
                    var res = await GoogleCalendarManager.Instance.Login(clientId, clientSecret);
                }
                reset();
            }
        }    

        private async void reset()
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
            trayIcon.ShowBalloonTip(title, text, Hardcodet.Wpf.TaskbarNotification.BalloonIcon.Info);
            HideBalloonAfterSeconds(6);
        }

        private bool[] colorChanged;
        private Hardcodet.Wpf.TaskbarNotification.TaskbarIcon trayIcon;
        private void textColorComboBox_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<System.Windows.Media.Color> e)
        {
            if(e.NewValue != e.OldValue)
            {
                colorChanged[0] = true;
            }
        }

        private void backgroundColorComboBox_SelectedColorChanged(object sender, RoutedPropertyChangedEventArgs<System.Windows.Media.Color> e)
        {
            if (e.NewValue != e.OldValue)
            {
                colorChanged[1] = true;
            }
        }

        private async void changeCalendarColor(System.Windows.Media.Color fg, System.Windows.Media.Color bg)
        {
            if(colorChanged[0] && colorChanged[1])
            {
                colorChanged[0] = false;
                colorChanged[1] = false;
                var foregroundColor =  "#" + fg.ToString().Substring(3); // remove alpha channel, we don't want that shit
                var backgroundColor = "#" + bg.ToString().Substring(3); 
                if(!GoogleCalendarManager.Instance.LoggedIn)
                {
                    var login = await GoogleCalendarManager.Instance.Login(clientId, clientSecret);
                }
                if(GoogleCalendarManager.Instance.LoggedIn)
                { 
                    try
                    {
                        var res = await GoogleCalendarManager.Instance.SetCalendarColor(backgroundColor.ToLower(), foregroundColor.ToLower());
                    }
                    catch(Exception ex)
                    {
                        throw ex;
                    }
                }
                trayIcon.ShowBalloonTip("Colore calendario", "Il colore del tuo calendario su google e' stato cambiato correttamente", Hardcodet.Wpf.TaskbarNotification.BalloonIcon.Info);
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
                trayIcon.HideBalloonTip();
            };
            tmr.Start();
        }

        private void textColorComboBox_LostFocus(object sender, RoutedEventArgs e)
        {
            changeCalendarColor(textColorComboBox.SelectedColor, backgroundColorComboBox.SelectedColor);
        }

        private void backgroundColorComboBox_LostFocus(object sender, RoutedEventArgs e)
        {
            changeCalendarColor(textColorComboBox.SelectedColor, backgroundColorComboBox.SelectedColor);
        }

        private void calnameTextBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            //todo: add summary customization here
        }
    }
}
