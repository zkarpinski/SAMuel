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
using System.Diagnostics;
using System.Threading;
using System.IO;

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
            if (Properties.Settings.Default.LogLocation == "DEFAULT")
            {
                Properties.Settings.Default.LogLocation = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                Properties.Settings.Default.Save();
            }
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
            dlg.ReadOnlyChecked = true;
            dlg.ShowDialog();

            //Stop if no files were selected.
            if (dlg.FileNames.Length <= 0) { System.Diagnostics.Debug.Print("No files selected to be faxed."); return; }
            else { files = dlg.FileNames; }

            BeginFaxProcess(files);

        }


        private void BeginFaxProcess(string[] files)
        {

            
            List<Fax> faxes;
            faxes = CreateFaxes(files: ref files, faxInfoType: FaxInfoType.Parsed);
            FileListBox.ItemsSource = faxes;

            for (int i = 0; i < faxes.Count; i++)
            {
                if (!faxes[i].IsValid)
                {
                    // TODO Invalid fax fixing
                    //Removes fax if invalid
                    faxes.RemoveAt(i);
                    i -= 1;
                }
            }
            LogFaxes(ref faxes);
            if (SendFaxes(ref faxes))
            {
                MessageBox.Show("Faxes sent!");
                // TODO Move fax
            }
        }

        /// <summary>
        /// Creates an array of faxes from a string of files.
        /// </summary>
        /// <param name="files">Array of files to be faxed.</param>
        /// <param name="faxInfoType"></param>
        /// <param name="recipient"></param>
        /// <param name="faxNumber"></param>
        /// <returns>Array of faxes</returns>
        private List<Fax> CreateFaxes(ref string[] files, FaxInfoType faxInfoType = FaxInfoType.Manual, string recipient = "DEFAULT", string faxNumber = "DEFAULT")
        {
            List<Fax> faxes = new List<Fax>();
            int i = 0;
            foreach (string file in files)
            {
                if (faxInfoType == FaxInfoType.Parsed)
                {
                    Fax fax = new Fax(file);
                    faxes.Add(fax);
                }
                else if (faxInfoType == FaxInfoType.Manual)
                {
                    Fax fax = new Fax(file, recipient, faxNumber);
                    faxes.Add(fax);
                }
                else
                {
                    Fax fax = new Fax();
                    faxes.Add(fax);
                }

                i += 1;
            }
            return faxes;

        }

        private Boolean SendFaxes(ref List<Fax> faxes)
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
                MessageBox.Show(System.Environment.NewLine + e, "RightFax Error");
                return false;
            }
        }



        private void MoveCompletedFax(string pathToDocument) 
        {
            String fileName = System.IO.Path.GetFileName(pathToDocument);
            String saveTo = @"‪C:\Users\zkarpinski\Desktop\fax\Complete\";
            try
            {
                System.IO.File.Move(pathToDocument, saveTo);
            }
            catch (Exception)
            {
                
                throw;
            }
            
        }

        private void OpenDocument(object sender, MouseButtonEventArgs e)
        {
            //FileListBox.SelectedItem;
        }

        private void LogFaxes(ref List<Fax> faxes)
        {
            String logFile = Properties.Settings.Default.LogLocation + "\\RightFax_It-log.txt";
            DateTime logTime = DateTime.Now;

            if (!File.Exists(logFile))
            {
                FileStream fs = File.Create(logFile);
                fs.Close();
            }

            File.AppendAllText(logFile, System.Environment.NewLine + logTime.ToString() + System.Environment.NewLine);

            foreach (Fax fax in faxes)
            {
                String logAction;
                logAction = fax.Account + "\t" + fax.CustomerName + "\t\t\t" + fax.FaxNumber + System.Environment.NewLine;
                File.AppendAllText(logFile, logAction);
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

    }
}
