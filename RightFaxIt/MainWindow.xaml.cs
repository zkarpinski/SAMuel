
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
using System.ComponentModel;
using System.Runtime.InteropServices;

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

            //Setups the the log location on first run unless the user changed it.
            if (Properties.Settings.Default.LogLocation == "DEFAULT")
            {
                Properties.Settings.Default.LogLocation = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                Properties.Settings.Default.Save();
            }
        }

        [Flags]
        private enum FaxInfoType { Parsed, Manual };
        public List<Fax> faxes;

        #region Background Worker
        private void bworker_DoWork(object sender, DoWorkEventArgs e)
        {

            String[] files = ((Tuple <string[],string>)e.Argument).Item1;
            String userID = ((Tuple<string[], string>)e.Argument).Item2;
            List<String> invalidFiles = new List<string>();
            
            //Create each fax object from the list of files.
            faxes = CreateFaxes(files: ref files, faxInfoType: FaxInfoType.Parsed);

            //Removes fax from list if invalid
            for (int i = 0; i < faxes.Count; i++)
            {
                if (!faxes[i].IsValid)
                {
                    // TODO Invalid fax fixing
                    invalidFiles.Add(faxes[i].Document);
                    faxes.RemoveAt(i);
                    i -= 1;
                }
            }

            //Exit if there are no valid faxes to send out.
            if (faxes.Count <= 0)
            {
                 e.Result = invalidFiles;
                return;
            }

            
            //Send all faxes out and move them if all were successful.
            if (SendFaxes(ref faxes,userID))
            {
                //Log all the faxed files.
                LogFaxes(ref faxes, userID);

                //Determine where to move the files.
                String moveFolder;
                switch (userID)
                {
                    case "active":
                        moveFolder = Properties.Settings.Default.ActiveMoveLocation;
                        break;
                    case "cutin":
                        moveFolder = Properties.Settings.Default.CutInMoveLocation;
                        break;
                    default:
                        moveFolder = Properties.Settings.Default.ActiveMoveLocation;
                    break;
                }
                
                //Move files when faxed successfully
                for (int i = 0; i < faxes.Count; i++)
                {
                    MoveCompletedFax(faxes[i].Document,moveFolder);
                }
            }
            else
            {
                //TODO Add failed send fax handling
            }

            //Return the skipped file list
            e.Result = invalidFiles;
            return;
        }

        private void bworker_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            //Retrieve the list of skipped files.
            List<string> skippedFiles = (List<string>)e.Result;
            string msg;

            //Display message of skipped files.
            if (skippedFiles.Count() > 0)
            {
                msg = skippedFiles.Count.ToString() + " document(s) skipped.";
                for(int i = 0; i<skippedFiles.Count();i++)
                {
                    msg = msg + Environment.NewLine + skippedFiles[i];
                }
                Console.WriteLine(msg);
                //String msgTitle = faxes.Count.ToString() + " items faxed.";
                //MessageBox.Show(msg, msgTitle);
            }

            //Update UI
            //this.AllowDrop = true;//update ui once worker complete his work
            //this.btnFax.IsEnabled = true;
            //FileListBox.ItemsSource = faxes;
        }

        void bworker_CompletedProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FaxProgress.Value = e.ProgressPercentage;
        }
        #endregion

        /// <summary>
        /// Creates a single fax.
        /// </summary>
        /// <param name="file"></param>
        /// <param name="faxInfoType"></param>
        /// <param name="recipient"></param>
        /// <param name="faxNumber"></param>
        /// <returns></returns>
        private Fax CreateFax(string file, FaxInfoType faxInfoType = FaxInfoType.Manual, string recipient = "DEFAULT", string faxNumber = "DEFAULT")
        {
            Fax fax;
                if (faxInfoType == FaxInfoType.Parsed)
                {
                    fax = new Fax(file);
                }
                else if (faxInfoType == FaxInfoType.Manual)
                {
                    fax = new Fax(file, recipient, faxNumber);
                }
                else
                {
                    fax = new Fax();
                }
            return fax;
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

        private Boolean SendFaxes(ref List<Fax> faxes, String userID)
        {
            try 
	        {	        
                //Setup Rightfax Server Connection
		        RFCOMAPILib.FaxServer faxsvr = new RFCOMAPILib.FaxServer();
                faxsvr.ServerName = Properties.Settings.Default.FaxServerName;
                faxsvr.AuthorizationUserID = userID;
                faxsvr.Protocol = RFCOMAPILib.CommunicationProtocolType.cpTCPIP;
                faxsvr.UseNTAuthentication = RFCOMAPILib.BoolType.False;
                faxsvr.OpenServer();

                //Create each fax and send.
                foreach (Fax fax in faxes)
                {
                    if (fax.IsValid)
                    {
                        RFCOMAPILib.Fax newFax = (RFCOMAPILib.Fax)faxsvr.get_CreateObject(RFCOMAPILib.CreateObjectType.coFax);
                        newFax.ToName = fax.CustomerName;
                        newFax.ToFaxNumber = Regex.Replace(fax.FaxNumber, "-", "");
                        newFax.Attachments.Add(fax.Document);
                        newFax.UserComments = "Sent via SAMuel.";
                        newFax.Send();
                        // TODO newFax.MoveToFolder
                    }
                    else
                    {  }
                }
                faxsvr.CloseServer();
            return true;
	        }

	        catch (Exception e)
	        {
                MessageBox.Show(System.Environment.NewLine + e, "RightFax Error");
                return false;
            }
        }



        private void MoveCompletedFax(string pathToDocument, string folderDestination) 
        {
            String fileName = System.IO.Path.GetFileName(pathToDocument);
            String saveTo = System.IO.Path.Combine(folderDestination, fileName);
            
            try
            {
                //Delete the file in the destination if it exists already.
                // since File.Move does not overwrite.
                if (System.IO.File.Exists(saveTo))
                {
                    System.IO.File.Delete(saveTo);
                }
                System.IO.File.Move(pathToDocument, saveTo);
            }
            catch (Exception)
            {
                throw;
            }
            
        }

        private void OpenDocument(object sender, MouseButtonEventArgs e)
        {
            Fax item = (Fax)FileListBox.SelectedItem;
            String docPath = item.Document;
            System.Diagnostics.Process.Start(docPath);            
        }

        private void LogFaxes(ref List<Fax> faxes,string userID)
        {
            String logFile = Properties.Settings.Default.LogLocation + "\\RightFax_It-log.txt";
            DateTime logTime = DateTime.Now;

            if (!File.Exists(logFile))
            {
                FileStream fs = File.Create(logFile);
                fs.Close();
            }

            File.AppendAllText(logFile, System.Environment.NewLine + logTime.ToString() + "\t" + userID + System.Environment.NewLine);

            foreach (Fax fax in faxes)
            {
                String logAction;
                logAction = fax.Account + "\t" + fax.CustomerName + "\t\t\t" + fax.FaxNumber + System.Environment.NewLine;
                File.AppendAllText(logFile, logAction);
            }
        }

        #region UI Interaction
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
            initFaxWorker(files);

        }

        private void FileDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                initFaxWorker(files);
            }
        }

        private void initFaxWorker(string[] files)
        {
            String SelectedUser;
            //this.AllowDrop = false;
            //this.btnFax.IsEnabled = false;
            
            if((bool)this.ActiveUserRatio.IsChecked) 
            { 
                SelectedUser = "active"; 
            }
            else if((bool)this.CutinUserRatio.IsChecked)
            {
                SelectedUser = "cutin";
            }
            else
            {
                SelectedUser = "active";
            }

            var arguments = Tuple.Create<string[], string>(files, SelectedUser);
            BackgroundWorker _backgroundWorker = new BackgroundWorker();
            _backgroundWorker.WorkerReportsProgress = true;
            _backgroundWorker.DoWork += bworker_DoWork;
            _backgroundWorker.RunWorkerCompleted += bworker_Completed;
            _backgroundWorker.RunWorkerAsync(arguments);
        }

        private void FileDrag(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Link;
            }
        }

        /// <summary>
        /// Closes the form.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuItem_Click(object sender, RoutedEventArgs e) { this.Close();}

        private void OptionsItem_Click(object sender, RoutedEventArgs e)
        {
            //var newWindow = new Options();
            //newWindow.Show();
        }

        private void AboutItem_Click(object sender, RoutedEventArgs e)
        {
            var newWindow = new About();
            newWindow.Show();
        }

#endregion

        /* Folder watching region of code --------------------------------------------------------*/
        #region Folder Watching

        /// <summary>
        /// Polls the desired folder.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void btnPoll_Click(object sender, RoutedEventArgs e)
        {
            StartQWorker();
            btnPoll.IsEnabled = false;
            //FileSystemWatcher fsw = new FileSystemWatcher(Properties.Settings.Default.PollingFolder);
            //fsw.NotifyFilter = NotifyFilters.CreationTime | NotifyFilters.FileName;
            //fsw.Created += new FileSystemEventHandler(FaxPolledFile);

            //fsw.EnableRaisingEvents = true;

           WatchFolder(Properties.Settings.Default.ActiveFolder, "active");
           WatchFolder(Properties.Settings.Default.CutinFolder, "cutin");
        }

        private void FaxPolledFile(object source, FileSystemEventArgs e)
        {
            String[] file = { e.FullPath };
            String sUser = "active";

            var arguments = Tuple.Create<string[], string>(file, sUser);
            BackgroundWorker _backgroundWorker = new BackgroundWorker();
            _backgroundWorker.WorkerReportsProgress = true;
            _backgroundWorker.DoWork += bworker_DoWork;
            _backgroundWorker.RunWorkerCompleted += bworker_Completed;
            _backgroundWorker.RunWorkerAsync(arguments);
        }

        private void WatchFolder(string sPath, string rightFaxUser)
        {
            // Create a new watcher for every folder we want to monitor.
            try
            {
            FileSystemWatcher fsw = new FileSystemWatcher(sPath, "*.doc");
            fsw.NotifyFilter = NotifyFilters.CreationTime | NotifyFilters.FileName;
            fsw.Created += (sender, e) => _NewFileCreated(sender,e,rightFaxUser);
            fsw.EnableRaisingEvents = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
                return;
            }
        }

        private void _NewFileCreated(object sender, FileSystemEventArgs e, String rightFaxUser)
        {
            Console.WriteLine("New File detected!");
            String file = e.FullPath;

            //Wait 2 seconds incase the file is being created still.
            Thread.Sleep(2000);
            
            //Create each fax object from the file.
            Fax fax = CreateFax(file,FaxInfoType.Parsed);

            //Skips the file if the fax is invalid.
             if (!fax.IsValid)
                {
                 fax = null;
                 return;
                }
             else
             {
                 Tuple<Fax, String> work = new Tuple<Fax, string>(fax, rightFaxUser);
                 AddFaxToQueue(work);
             }
            }
        #endregion


        #region FaxingQueue
        /// <summary>
        /// http://social.msdn.microsoft.com/forums/vstudio/en-US/500cb664-e2ca-4d76-88b9-0faab7e7c443/queuing-backgroundworker-tasks
        /// </summary>

        private Thread QueueWorker;
        private Object zLock = new object();
        private Queue<Tuple<Fax, String>> FaxQueue = new Queue<Tuple<Fax, String>>(50);
        private Boolean  QuitWork;
        private EventWaitHandle DoQWork = new EventWaitHandle(false,EventResetMode.ManualReset);

        private void StopQWorker()
        {
            QuitWork = true;
            DoQWork.Set();
            QueueWorker.Join(1000);
        }

        private void StartQWorker()
        {
            QueueWorker = new Thread(QThread);
            QueueWorker.IsBackground = true;
            QueueWorker.Start();
        }

        private void AddFaxToQueue (Tuple<Fax,String> work )
        {
            lock(zLock)
            {
                FaxQueue.Enqueue(work);
            }
            DoQWork.Set();
        }

        private void QThread()
        {
            Console.WriteLine("Thread Started.");
          do
          {
              Console.WriteLine("Thread Waiting.");
              DoQWork.WaitOne(-1,false);
              Console.WriteLine("Checking for work.");
              if (QuitWork) { break;}
              Tuple<Fax,String> dequeuedWork;
              do
              {
                  dequeuedWork = null;
                  Console.WriteLine("Dequeueing");
                  lock(zLock)
                  {
                      if (FaxQueue.Count > 0)
                      {
                          dequeuedWork = FaxQueue.Dequeue();
                      }
                  }

                  if (dequeuedWork != null)
                  {
                      Console.WriteLine("Working");

                      ProcessFax(dequeuedWork);
                      Console.WriteLine("Work Completed!");
                  }
              }while(dequeuedWork != null);

              lock(zLock)
              {
                  if (FaxQueue.Count == 0)
                  {
                      DoQWork.Reset();
                  }
              }

          }while(true);
          Console.WriteLine("THREAD ENDED");
        }

        #endregion

        private void ProcessFax(Tuple<Fax,String> work)
        {
            if (SendFax(work.Item1, work.Item2))
            {
                //Determine where to move the files.
                String moveFolder;
                switch (work.Item2)
                {
                    case "active":
                        moveFolder = Properties.Settings.Default.ActiveMoveLocation;
                        break;
                    case "cutin":
                        moveFolder = Properties.Settings.Default.CutInMoveLocation;
                        break;
                    default:
                        moveFolder = Properties.Settings.Default.ActiveMoveLocation;
                    break;
                }
                MoveCompletedFax(work.Item1.Document, moveFolder);
            }   
        }

        private bool SendFax(Fax fax, String userID)
        {
            try
            {
                //Setup Rightfax Server Connection
                RFCOMAPILib.FaxServer faxsvr = new RFCOMAPILib.FaxServer();
                faxsvr.ServerName = Properties.Settings.Default.FaxServerName;
                faxsvr.AuthorizationUserID = userID;
                faxsvr.Protocol = RFCOMAPILib.CommunicationProtocolType.cpTCPIP;
                faxsvr.UseNTAuthentication = RFCOMAPILib.BoolType.False;
                faxsvr.OpenServer();

                //Create the fax and send.
                if (fax.IsValid)
                    {
                        RFCOMAPILib.Fax newFax = (RFCOMAPILib.Fax)faxsvr.get_CreateObject(RFCOMAPILib.CreateObjectType.coFax);
                        newFax.ToName = fax.CustomerName;
                        newFax.ToFaxNumber = Regex.Replace(fax.FaxNumber, "-", "");
                        newFax.Attachments.Add(fax.Document);
                        newFax.UserComments = "Sent via SAMuel.";
                        newFax.Send();
                        // TODO newFax.MoveToFolder
                    }
                    else
                    { return false; }
                faxsvr.CloseServer();
                return true;
            }

            catch (Exception e)
            {
                MessageBox.Show(System.Environment.NewLine + e, "RightFax Error");
                return false;
            }
        }


    }
}

