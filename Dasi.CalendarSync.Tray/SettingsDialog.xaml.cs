using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace Dasi.CalendarSync.Tray
{
    /// <summary>
    /// Logica di interazione per SettingsDialog.xaml
    /// </summary>
    public partial class SettingsDialog : Window
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
    }
}
