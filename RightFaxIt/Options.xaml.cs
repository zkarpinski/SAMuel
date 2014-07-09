using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace RightFaxIt
{
    /// <summary>
    /// Interaction logic for Options.xaml
    /// </summary>
    public partial class Options : Window
    {
        public Options()
        {

            InitializeComponent();
            txtUserID.Text = Properties.Settings.Default.FaxUserID;
            txtServerName.Text = Properties.Settings.Default.FaxServerName;
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            Properties.Settings.Default.FaxUserID = txtUserID.Text;
            Properties.Settings.Default.FaxServerName = txtServerName.Text;
            Properties.Settings.Default.Save();
            this.Close();
        }
    }
}
