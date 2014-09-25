using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows;
using System.Windows.Input;
using System.Windows.Interop;
using Microsoft.Win32;
using RFCOMAPILib;
using RightFaxIt.Properties;

namespace RightFaxIt
{
    /// <summary>
    ///     Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public List<Fax> Faxes;

        #region Background Worker

        private void bworker_DoWork(object sender, DoWorkEventArgs e)
        {
            var files = ((Tuple<string[], string>) e.Argument).Item1;
            var userId = ((Tuple<string[], string>) e.Argument).Item2;
            var invalidFiles = new List<string>();

            //Create each fax object from the list of files.
            Faxes = CreateFaxes(ref files, FaxInfoType.Parsed);

            //Removes fax from list if invalid
            for (var i = 0; i < Faxes.Count; i++)
            {
                if (!Faxes[i].IsValid)
                {
                    // TODO Invalid fax fixing
                    invalidFiles.Add(Faxes[i].Document);
                    Faxes.RemoveAt(i);
                    i -= 1;
                }
            }

            //Exit if there are no valid faxes to send out.
            if (Faxes.Count <= 0)
            {
                e.Result = invalidFiles;
                return;
            }


            //Send all faxes out and move them if all were successful.
            if (SendFaxes(ref Faxes, userId))
            {
                //Log all the faxed files.
                LogFaxes(ref Faxes, userId);

                //Determine where to move the files.
                String moveFolder;
                switch (userId)
                {
                    case "active":
                        moveFolder = Settings.Default.ActiveMoveLocation;
                        break;
                    case "cutin":
                        moveFolder = Settings.Default.CutInMoveLocation;
                        break;
                    default:
                        moveFolder = Settings.Default.ActiveMoveLocation;
                        break;
                }

                //Move files when faxed successfully
                foreach (var fax in Faxes)
                {
                    try
                    {
                        MoveCompletedFax(fax.Document, moveFolder);
                    }
                    catch (Exception ex)
                    {
                        //TODO Add log action
                    }

                }
            }

            //Return the skipped file list
            e.Result = invalidFiles;
        }

        private void bworker_Completed(object sender, RunWorkerCompletedEventArgs e)
        {
            
            //TODO Added completed actions.
            //Retrieve the list of skipped files.
            var skippedFiles = (List<string>) e.Result;

            //Display message of skipped files.
            if (!skippedFiles.Any()) return;

            var msg = skippedFiles.Count + " document(s) skipped.";
            for (var i = 0; i < skippedFiles.Count(); i++)
            {
                msg = msg + Environment.NewLine + skippedFiles[i];
            }
            Console.WriteLine(msg);

        }

        private void bworker_CompletedProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FaxProgress.Value = e.ProgressPercentage;
        }

        #endregion

        public MainWindow()
        {
            InitializeComponent();
            //Setups the the log location on first run unless the user changed it.
            if (Settings.Default.LogLocation == "DEFAULT")
            {
                Settings.Default.LogLocation = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                Settings.Default.Save();
            }
        }

        /// <summary>
        ///     Creates a single fax.
        /// </summary>
        /// <param name="file"></param>
        /// <param name="faxInfoType"></param>
        /// <param name="recipient"></param>
        /// <param name="faxNumber"></param>
        /// <returns></returns>
        private Fax CreateFax(string file, FaxInfoType faxInfoType = FaxInfoType.Manual, string recipient = "DEFAULT",
            string faxNumber = "DEFAULT")
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
        ///     Creates an array of faxes from a string of files.
        /// </summary>
        /// <param name="files">Array of files to be faxed.</param>
        /// <param name="faxInfoType"></param>
        /// <param name="recipient"></param>
        /// <param name="faxNumber"></param>
        /// <returns>Array of faxes</returns>
        private List<Fax> CreateFaxes(ref string[] files, FaxInfoType faxInfoType = FaxInfoType.Manual,
            string recipient = "DEFAULT", string faxNumber = "DEFAULT")
        {
            var faxes = new List<Fax>();
            var i = 0;
            foreach (var file in files)
            {
                if (faxInfoType == FaxInfoType.Parsed)
                {
                    var fax = new Fax(file);
                    faxes.Add(fax);
                }
                else if (faxInfoType == FaxInfoType.Manual)
                {
                    var fax = new Fax(file, recipient, faxNumber);
                    faxes.Add(fax);
                }
                else
                {
                    var fax = new Fax();
                    faxes.Add(fax);
                }

                i += 1;
            }
            return faxes;
        }

        private Boolean SendFaxes(ref List<Fax> faxes, String userId)
        {
            try
            {
                //Setup Rightfax Server Connection
                var faxsvr = new FaxServer
                {
                    ServerName = Settings.Default.FaxServerName,
                    AuthorizationUserID = userId,
                    Protocol = CommunicationProtocolType.cpTCPIP,
                    UseNTAuthentication = BoolType.False
                };
                faxsvr.OpenServer();

                //Create each fax and send.
                foreach (var fax in faxes)
                {
                    if (fax.IsValid)
                    {
                        var newFax = (RFCOMAPILib.Fax) faxsvr.get_CreateObject(CreateObjectType.coFax);
                        newFax.ToName = fax.CustomerName;
                        newFax.ToFaxNumber = Regex.Replace(fax.FaxNumber, "-", "");
                        newFax.Attachments.Add(fax.Document);
                        newFax.UserComments = "Sent via SAMuel.";
                        newFax.Send();
                        // TODO newFax.MoveToFolder
                    }
                }
                faxsvr.CloseServer();
                return true;
            }

            catch (Exception e)
            {
                MessageBox.Show(Environment.NewLine + e, "RightFax Error");
                return false;
            }
        }


        private void MoveCompletedFax(string pathToDocument, string folderDestination)
        {
            var fileName = Path.GetFileName(pathToDocument);
            if (fileName == null) return;
            var saveTo = Path.Combine(folderDestination, fileName);

            //Delete the file in the destination if it exists already.
            // since File.Move does not overwrite.
            if (File.Exists(saveTo))
            {
                File.Delete(saveTo);
            }
            File.Move(pathToDocument, saveTo);
        }

        private static void LogFaxes(ref List<Fax> faxes, string userId)
        {
            var logFile = Settings.Default.LogLocation + "\\RightFax_It-log.txt";
            var logTime = DateTime.Now;

            if (!File.Exists(logFile))
            {
                var fs = File.Create(logFile);
                fs.Close();
            }

            File.AppendAllText(logFile, Environment.NewLine + logTime + "\t" + userId + Environment.NewLine);

            foreach (var fax in faxes)
            {
                var logAction = fax.Account + "\t" + fax.CustomerName + "\t\t\t" + fax.FaxNumber + Environment.NewLine;
                File.AppendAllText(logFile, logAction);
            }
        }

        /* Folder watching region of code --------------------------------------------------------*/

        private void ProcessFax(Tuple<Fax, String> work)
        {
            if (SendFax(work.Item1, work.Item2))
            {
                //Determine where to move the files.
                String moveFolder;
                switch (work.Item2)
                {
                    case "active":
                        moveFolder = Settings.Default.ActiveMoveLocation;
                        break;
                    case "cutin":
                        moveFolder = Settings.Default.CutInMoveLocation;
                        break;
                    default:
                        moveFolder = Settings.Default.ActiveMoveLocation;
                        break;
                }
                MoveCompletedFax(work.Item1.Document, moveFolder);
                LogFax(work.Item1, work.Item2);
            }
        }

        private void LogFax(Fax fax, string userId)
        {
            var logFile = Settings.Default.LogLocation + "\\RightFax_It-log.txt";
            var logTime = DateTime.Now;

            if (!File.Exists(logFile))
            {
                var fs = File.Create(logFile);
                fs.Close();
            }

            var logAction = logTime + " \t AUTOFAX: " + userId + "\t" + fax.Account + "\t" + fax.FaxNumber + "\t" +
                               fax.CustomerName + Environment.NewLine;

            File.AppendAllText(logFile, logAction);
        }

        private bool SendFax(Fax fax, String userId)
        {
            try
            {
                //Setup Rightfax Server Connection
                var faxsvr = new FaxServer
                {
                    ServerName = Settings.Default.FaxServerName,
                    AuthorizationUserID = userId,
                    Protocol = CommunicationProtocolType.cpTCPIP,
                    UseNTAuthentication = BoolType.False
                };
                faxsvr.OpenServer();

                //Create the fax and send.
                if (fax.IsValid)
                {
                    var newFax = (RFCOMAPILib.Fax) faxsvr.get_CreateObject(CreateObjectType.coFax);
                    newFax.ToName = fax.CustomerName;
                    newFax.ToFaxNumber = Regex.Replace(fax.FaxNumber, "-", "");
                    newFax.Attachments.Add(fax.Document);
                    newFax.UserComments = "Sent via SAMuel.";
                    newFax.Send();
                    // TODO newFax.MoveToFolder
                }
                else
                {
                    return false;
                }
                faxsvr.CloseServer();
                return true;
            }

            catch (Exception e)
            {
                MessageBox.Show(Environment.NewLine + e, "RightFax Error");
                return false;
            }
        }

        #region UI Interaction

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            //Setup file dialog box
            var dlg = new OpenFileDialog
            {
                Multiselect = true,
                Filter = "Word Documents|*.doc;*.docx|All files (*.*)|*.*",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments),
                Title = "Select the documents you wish to fax.",
                ReadOnlyChecked = true
            };
            dlg.ShowDialog();

            //Stop if no files were selected.
            if (dlg.FileNames.Length <= 0)
            {
                Debug.Print("No files selected to be faxed.");
                return;
            }
            var files = dlg.FileNames;
            InitFaxWorker(files);
        }

        private void FileDrop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                var files = (string[]) e.Data.GetData(DataFormats.FileDrop);
                InitFaxWorker(files);
            }
        }

        private void InitFaxWorker(string[] files)
        {
            String selectedUser;
            //this.AllowDrop = false;
            //this.btnFax.IsEnabled = false;

            if ((bool) ActiveUserRatio.IsChecked)
            {
                selectedUser = "active";
            }
            else if ((bool) CutinUserRatio.IsChecked)
            {
                selectedUser = "cutin";
            }
            else
            {
                selectedUser = "active";
            }

            var arguments = Tuple.Create(files, selectedUser);
            var backgroundWorker = new BackgroundWorker {WorkerReportsProgress = true};
            backgroundWorker.DoWork += bworker_DoWork;
            backgroundWorker.RunWorkerCompleted += bworker_Completed;
            backgroundWorker.RunWorkerAsync(arguments);
        }

        private void FileDrag(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Link;
            }
        }

        /// <summary>
        ///     Closes the form.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void MenuItem_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }

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

        #region Folder Watching

        /// <summary>
        ///     Polls the desired folder.
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

            WatchFolder(Settings.Default.ActiveFolder, "active");
            WatchFolder(Settings.Default.CutinFolder, "cutin");
        }

        //private void FaxPolledFile(object source, FileSystemEventArgs e)
        //{
        //    String[] file = {e.FullPath};
        //    String sUser = "active";

        //    Tuple<string[], string> arguments = Tuple.Create(file, sUser);
        //    var _backgroundWorker = new BackgroundWorker();
        //    _backgroundWorker.WorkerReportsProgress = true;
        //    _backgroundWorker.DoWork += bworker_DoWork;
        //    _backgroundWorker.RunWorkerCompleted += bworker_Completed;
        //    _backgroundWorker.RunWorkerAsync(arguments);
        //}

        private void WatchFolder(string sPath, string rightFaxUser)
        {
            // Create a new watcher for every folder we want to monitor.
            try
            {
                var fsw = new FileSystemWatcher(sPath, "*.doc")
                {
                    NotifyFilter = NotifyFilters.CreationTime | NotifyFilters.FileName
                };
                fsw.Created += (sender, e) => _NewFileCreated(sender, e, rightFaxUser);
                fsw.EnableRaisingEvents = true;
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
        }

        private void _NewFileCreated(object sender, FileSystemEventArgs e, String rightFaxUser)
        {
            Console.WriteLine("New File detected!");
            var file = e.FullPath;

            //Wait 2 seconds incase the file is being created still.
            Thread.Sleep(2000);

            //Create each fax object from the file.
            var fax = CreateFax(file, FaxInfoType.Parsed);

            //Skips the file if the fax is invalid.
            if (!fax.IsValid)
            {
                fax = null;
            }
            var work = new Tuple<Fax, string>(fax, rightFaxUser);
            AddFaxToQueue(work);
        }

        #endregion

        #region FaxingQueue

        private readonly EventWaitHandle _doQWork = new EventWaitHandle(false, EventResetMode.ManualReset);
        private readonly Queue<Tuple<Fax, String>> _faxQueue = new Queue<Tuple<Fax, String>>(50);
        private readonly Object _zLock = new object();

        /// <summary>
        ///     http://social.msdn.microsoft.com/forums/vstudio/en-US/500cb664-e2ca-4d76-88b9-0faab7e7c443/queuing-backgroundworker-tasks
        /// </summary>
        private Thread _queueWorker;

        private Boolean _quitWork;

        private void StopQWorker()
        {
            _quitWork = true;
            _doQWork.Set();
            _queueWorker.Join(1000);
        }

        private void StartQWorker()
        {
            _queueWorker = new Thread(QThread) {IsBackground = true};
            _queueWorker.Start();
        }

        private void AddFaxToQueue(Tuple<Fax, String> work)
        {
            lock (_zLock)
            {
                _faxQueue.Enqueue(work);
            }
            _doQWork.Set();
        }

        private void QThread()
        {
            Console.WriteLine("Thread Started.");
            do
            {
                Console.WriteLine("Thread Waiting.");
                _doQWork.WaitOne(-1, false);
                Console.WriteLine("Checking for work.");
                if (_quitWork)
                {
                    break;
                }
                Tuple<Fax, String> dequeuedWork;
                do
                {
                    dequeuedWork = null;
                    Console.WriteLine("Dequeueing");
                    lock (_zLock)
                    {
                        if (_faxQueue.Count > 0)
                        {
                            dequeuedWork = _faxQueue.Dequeue();
                        }
                    }

                    if (dequeuedWork != null)
                    {
                        Console.WriteLine("Working");

                        ProcessFax(dequeuedWork);
                        Console.WriteLine("Work Completed!");
                    }
                } while (dequeuedWork != null);

                lock (_zLock)
                {
                    if (_faxQueue.Count == 0)
                    {
                        _doQWork.Reset();
                    }
                }
            } while (true);
            Console.WriteLine("THREAD ENDED");
        }

        #endregion

        [Flags]
        private enum FaxInfoType
        {
            Parsed,
            Manual
        };
    }
}