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
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Text.RegularExpressions;

namespace RightFaxIt
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        [Flags]
        private enum FaxInfoType { Parsed, Manual };

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            String[] files;

            //Setup file dialog box
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.Multiselect = true;
            dlg.Filter = "Word Documents|*.doc;*.docx|All files (*.*)|*.*";
            dlg.InitialDirectory = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
            dlg.Title = "Select the documents you wish to fax.";
            dlg.ShowDialog();

            if (dlg.FileNames.Length < 0)
            {
                System.Diagnostics.Debug.Print ("No files selected to be faxed.");
                return;
            }

            files = dlg.FileNames;

            BeginFaxProcess(files);


        }

        private void BeginFaxProcess(string[] files)
        {

            Fax[] faxes;
            faxes = CreateFaxes(files: ref files, faxInfoType: FaxInfoType.Parsed);
            FileListBox.ItemsSource = faxes;

            if (SendFaxes(ref faxes))
            {
                MessageBox.Show("Faxes sent!");
            }
        }


        private void FileDrag(object sender, DragEventArgs e)
        {
                if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
        }

        private void FileDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                BeginFaxProcess(files);
            }
        }

        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

        private void OptionsItem_Click(object sender, RoutedEventArgs e)
        {
            var newWindow = new Options();
            newWindow.Show();
        }

        private Boolean SendFaxes(ref Fax[] faxes)
        {
            try 
	        {	        
		        RFCOMAPILib.FaxServer faxsvr = new RFCOMAPILib.FaxServer();
                faxsvr.ServerName = Properties.Settings.Default.FaxServerName;
                faxsvr.AuthorizationUserID = Properties.Settings.Default.FaxUserID;
                faxsvr.Protocol = RFCOMAPILib.CommunicationProtocolType.cpTCPIP;
                faxsvr.UseNTAuthentication = RFCOMAPILib.BoolType.False;
                faxsvr.OpenServer();

                foreach (Fax fax in faxes)
                {
                    if (fax.IsValid)
                    {
                        RFCOMAPILib.Fax newFax = (RFCOMAPILib.Fax)faxsvr.get_CreateObject(RFCOMAPILib.CreateObjectType.coFax);
                        newFax.ToName = fax.CustomerName;
                        newFax.ToFaxNumber = Regex.Replace(fax.FaxNumber, "-", "");
                        newFax.Attachments.Add(fax.Document);
                        newFax.Send();
                    }
                    else
                        MessageBox.Show(fax.FileName + " is invalid. Skipping...");
                }
            return true;
	        }

	        catch (Exception e)
	        {
                MessageBox.Show("RightFax Error: " + e);
                return false;
            }
        }

        /// <summary>
        /// Creates an array of faxes from a string of files.
        /// </summary>
        /// <param name="files">Array of files to be faxed.</param>
        /// <param name="faxInfoType"></param>
        /// <param name="recipient"></param>
        /// <param name="faxNumber"></param>
        /// <returns></returns>
        private Fax[] CreateFaxes(ref string[] files, FaxInfoType faxInfoType = FaxInfoType.Manual, string recipient = "DEFAULT", string faxNumber = "DEFAULT")
        {
            Fax[] faxes = new Fax[files.Length];
            int i = 0;
            foreach (string file in files)
            {
                if (faxInfoType == FaxInfoType.Parsed)
                {
                    Fax fax = new Fax(file);
                    faxes[i] = fax;
                }
                else if (faxInfoType == FaxInfoType.Manual)
                {
                    Fax fax = new Fax(file, recipient, faxNumber);
                    faxes[i] = fax;
                }
                else
                {
                    Fax fax = new Fax();
                    faxes[i] = fax;
                }

                i += 1;
            }
            return faxes;

        }


    }
}
