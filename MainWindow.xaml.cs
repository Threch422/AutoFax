using System;
using System.Collections.Specialized;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Text;
using System.Configuration;
using System.Windows;
using System.Windows.Forms;

namespace AutoFax
{
    /// <summary>
    /// MainWindow.xaml 的互動邏輯
    /// </summary>
    public partial class MainWindow : Window
    {
        private NameValueCollection appSettings;

        protected string selectedFolder = string.Empty;
        protected string selectedExcel = string.Empty;

        BackgroundWorker worker = new BackgroundWorker();

        public MainWindow()
        {
            InitializeComponent();

            this.LoadAppSettings();
            // Redirect the console output to ProcessLogTextBox
            Console.SetOut(new ControlWriter(this.ProcessLogTextBox));

            this.worker.WorkerReportsProgress = true;
            this.worker.WorkerSupportsCancellation = true;
        }

        private void MainWindowClosing(object sender, CancelEventArgs e)
        {
            if (this.worker.IsBusy)
            {
                worker.CancelAsync();
                e.Cancel = true;
                System.Windows.Forms.MessageBox.Show($"The closing operation is accepted. Please wait for the current fax process ends before closing the window.", "Fax Process", MessageBoxButtons.OK);
                return;
            }
        }

        private void LoadAppSettings()
        {
            try
            {
                this.appSettings = ConfigurationManager.AppSettings;
            }
            catch (Exception e)
            {
                var result = System.Windows.Forms.MessageBox.Show(e.Message, "Error Detected", MessageBoxButtons.OK);
                this.Close();
            }
        }

        private void DirButton_Click(object sender, RoutedEventArgs e)
        {
            using (FolderBrowserDialog dialog = new FolderBrowserDialog())
            {
                dialog.ShowNewFolderButton = false;
                dialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

                var result = dialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    this.selectedFolder = dialog.SelectedPath;
                    this.DirTextBox.Text = this.selectedFolder;
                }

                // Check if variables selectedFolder and selected Excel are both set
                // If they are set, enable the Send Fax Button
                this.EnableSendFaxButton();
            }
        }

        private void ExcelFileButton_Click(object sender, RoutedEventArgs e)
        {
            using (OpenFileDialog openFileDialog = new OpenFileDialog())
            {
                openFileDialog.InitialDirectory = this.appSettings["DefaultExcelFilePath"];
                openFileDialog.Filter = "Excel File (*.xlsx)|*.xlsx|Excel File 97-2003 (*.xls)|*.xls";
                openFileDialog.FilterIndex = 2;
                openFileDialog.RestoreDirectory = true;

                var result = openFileDialog.ShowDialog();

                if (result == System.Windows.Forms.DialogResult.OK)
                {
                    this.selectedExcel = openFileDialog.FileName;
                    this.ExcelFileTextBox.Text = this.selectedExcel;
                }

                // Check if variables selectedFolder and selected Excel are both set
                // If they are set, enable the Send Fax Button
                this.EnableSendFaxButton();
            }

        }

        internal void EnableSendFaxButton()
        {
            if (!string.IsNullOrEmpty(this.selectedExcel) && !string.IsNullOrEmpty(this.selectedExcel))
            {
                this.SendFaxButton.IsEnabled = true;
            }
            else
            {
                this.SendFaxButton.IsEnabled = false;
            }
        }

        private async void SendFaxButton_Click(object sender, RoutedEventArgs e)
        {
            if (!Directory.Exists("./Log"))
                Directory.CreateDirectory("./Log");

            var excelHandler = new ExcelHandler(this.selectedExcel);

            // Filter out unknown and invisible temporary file
            var allowedExtension = this.appSettings["AcceptFileExtension"].Split(new char[] { ';' }).ToList();
            var documents = Directory.EnumerateFiles(this.selectedFolder, "*.*", SearchOption.AllDirectories)
                                .Where(fullpath => !Path.GetFileNameWithoutExtension(fullpath).StartsWith("~$") && allowedExtension.Any(fullpath.ToLower().EndsWith))
                                .ToDictionary(fullpath => Path.GetFileNameWithoutExtension(fullpath),
                                              fullpath => fullpath);

            foreach ((var faxNum, var recipientName) in excelHandler.GetRowsInfo())
                Console.WriteLine($"{faxNum}, {recipientName}");

            worker.DoWork += (workerSender, workerEvent) =>
            {
                for (int i = 0; i < 100; i++)
                {
                    if (worker.CancellationPending == true)
                    {
                        workerEvent.Cancel = true;
                        System.Windows.Forms.MessageBox.Show($"The fax process has ended. You can now close the window :).", "Fax Process", MessageBoxButtons.OK);
                        break;
                    }
                    System.Threading.Thread.Sleep(1000);
                    worker.ReportProgress(i + 1);
                }
            };

            worker.ProgressChanged += (workerSender, workerEvent) =>
            {
                ProgressBar.Value += 1;

                Console.WriteLine("Updated");
            };

            worker.RunWorkerCompleted += (workerSender, workerEvent) =>
            {
            };

            worker.RunWorkerAsync();


            //var faxSender = new CustomFaxSender();

            //foreach ((string faxNumber, string recipientName) in excelHandler.GetRowsInfo())
            //{
            //    if (documents.TryGetValue(recipientName, out var docPath))
            //    {
            //        var recipientInfo = new Dictionary<string, string> { { faxNumber, string.Empty } };
            //        faxSender.SendFax(recipientInfo, documents[recipientName]);
            //    }
            //}

            //faxSender.DisconnectFaxServer();

            // Write the log file
            //using (StreamWriter writer = new StreamWriter($"./Log/FaxLog_{DateTime.Now.ToString("yyyyMMddHHmmss")}.txt"))
            //{
            //    writer.WriteLine(this.ProcessLogTextBox.Text);
            //}


        }
    }

    public class ControlWriter : TextWriter
    {
        private System.Windows.Controls.TextBox textbox;
        public ControlWriter(System.Windows.Controls.TextBox textbox)
        {
            this.textbox = textbox;
        }

        public override void Write(char value)
        {
            textbox.Text += value;
        }

        public override void Write(string value)
        {
            textbox.Text += value;
        }

        public override Encoding Encoding
        {
            get { return Encoding.ASCII; }
        }
    }
}
