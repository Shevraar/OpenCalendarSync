using System.Windows;
using System.Windows.Input;

namespace Dasi.CalendarSync.Tray
{
    /// <summary>
    /// Logica di interazione per SettingsDialog.xaml
    /// </summary>
    public partial class SettingsDialog
    {
        public SettingsDialog()
        {
            InitializeComponent();
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

        private void btReset_Click(object sender, RoutedEventArgs e)
        {
            Reset = true; //todo: ripensare logica reset
        }

        public bool Reset = false;
    }
}
