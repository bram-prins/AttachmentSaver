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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace AttachmentSaver
{
    /// <summary>
    /// Interaction logic for RunDatepicker.xaml
    /// </summary>
    public partial class RunDatepicker : Window
    {
        public RunDatepicker()
        {
            InitializeComponent();
        }

        private void confirm_Click(object sender, RoutedEventArgs e)
        {
            if (datePicker.SelectedDate != null)
            {
                var settings = Properties.Settings.Default;
                settings.lastRun = (DateTime)datePicker.SelectedDate;
                settings.Save();
                Globals.Ribbons._Ribbon.setLastRunBtn.Label = $"Laatst geïmporteerd om: {datePicker.SelectedDate}";
            }
            Hide();
        }

        private void cancel_Click(object sender, RoutedEventArgs e)
        {
            Hide();
        }
    }
}
