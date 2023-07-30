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
                System.Windows.Forms.MessageBox.Show($"The closing operation can be performed by pressing the OK button. However, please wait for the current fax process to finish before closing the window.", "Terminating Fax Process", MessageBoxButtons.OK, MessageBoxIcon.Information);
                worker.CancelAsync();
                e.Cancel = true;

                return;
            }
        }

        private void MainWindowClosed(object sender, EventArgs e)
        {
            System.Windows.Application.Current.Shutdown();
        }

        private void LoadAppSettings()
        {
            try
            {
                this.appSettings = ConfigurationManager.AppSettings;
            }
            catch (Exception e)
            {
                System.Windows.Forms.MessageBox.Show(e.Message, "App Setting Loading Error", MessageBoxButtons.OK);
                this.Close();
            }
        }

        private void DirButton_Click(object sender, RoutedEventArgs e)
        {
            if (IsGenerateWordDoc_CheckBox.IsChecked ?? false)
            {
                using (OpenFileDialog openFileDialog = new OpenFileDialog())
                {
                    openFileDialog.InitialDirectory = this.appSettings["DefaultExcelFilePath"];
                    openFileDialog.Filter = "Word Document (*.docx)|*.docx|Word Document 97-2003 (*.doc)|*.doc";
                    openFileDialog.FilterIndex = 2;
                    openFileDialog.RestoreDirectory = true;

                    var result = openFileDialog.ShowDialog();

                    if (result == System.Windows.Forms.DialogResult.OK)
                    {
                        this.selectedFolder = openFileDialog.FileName;
                        this.DirTextBox.Text = this.selectedFolder;
                    }

                    // Check if variables selectedFolder and selected Excel are both set
                    // If they are set, enable the Send Fax Button
                    this.EnableSendFaxButton();
                }
            }
            else
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
                }
            }

            // Check if variables selectedFolder and selected Excel are both set
            // If they are set, enable the Send Fax Button
            this.EnableSendFaxButton();
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
            if (!string.IsNullOrEmpty(this.selectedFolder) && !string.IsNullOrEmpty(this.selectedExcel))
            {
                this.SendFaxButton.IsEnabled = true;
            }
            else
            {
                this.SendFaxButton.IsEnabled = false;
            }
        }

        private void SendFaxButton_Click(object sender, RoutedEventArgs e)
        {
            // Checking whether the user wants to generate word documents or send fax
            if (IsGenerateWordDoc_CheckBox.IsChecked ?? false)
            {
                // Create a directory to save the generated word documents
                if (!Directory.Exists("./wordResult"))
                    Directory.CreateDirectory("./wordResult");

                var excelHandler = new ExcelHandler(this.selectedExcel);
                var wordHandler = new WordHandler(this.selectedFolder);

                try
                {
                    var replacementPairDict = excelHandler.GetTemplateReplacementList();
                    wordHandler.GenerateDocument(replacementPairDict);

                    // Choose Save Location
                    var chosenSaveDir = string.Empty;

                    using (FolderBrowserDialog dialog = new FolderBrowserDialog())
                    {
                        dialog.ShowNewFolderButton = false;
                        dialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);

                        var result = dialog.ShowDialog();

                        if (result == System.Windows.Forms.DialogResult.OK)
                        {
                            chosenSaveDir = dialog.SelectedPath;
                        }
                    }

                    // Move the files to user defined location, and delete the files in original directory
                    var files = Directory.GetFiles($"{Directory.GetCurrentDirectory()}/wordResult/");

                    foreach (var file in files)
                    {
                        File.Copy(file, $"{chosenSaveDir}//{Path.GetFileName(file)}", true);
                        File.Delete(file);
                    }

                    // Pops a message box to notify user the process is completed
                    System.Windows.Forms.MessageBox.Show("The word documents are generated :)", "Generate Documents", MessageBoxButtons.OK);
                }
                catch (Exception ex)
                {
                    Console.Write(ex);
                };

            }
            else
            {
                // Disable the Send Fax Button
                this.SendFaxButton.IsEnabled = false;

                // Clear the ProcessLogTextBox
                this.ProcessLogTextBox.Text = string.Empty;

                this.ProgressBar.Value = 0;

                // Check whether the Log Folder exists
                if (!Directory.Exists("./Log"))
                    Directory.CreateDirectory("./Log");

                var excelHandler = new ExcelHandler(this.selectedExcel);
                var documents = new Dictionary<string, string>();

                // Filter out unknown and invisible temporary file
                var allowedExtension = this.appSettings["AcceptFileExtension"].Split(new char[] { ';' }).ToList();

                try
                {
                    documents = Directory.EnumerateFiles(this.selectedFolder, "*.*", SearchOption.AllDirectories)
                                        .Where(fullpath => !Path.GetFileNameWithoutExtension(fullpath).StartsWith("~$") && allowedExtension.Any(fullpath.ToLower().EndsWith))
                                        .Distinct()
                                        .ToDictionary(fullpath => Path.GetFileNameWithoutExtension(fullpath),
                                                      fullpath => fullpath);
                }
                catch (Exception ex)
                {
                    System.Windows.Forms.MessageBox.Show($"{ex.Message}", "Duplicate File Name", MessageBoxButtons.OK);

                    return;
                }

                var faxSender = new CustomFaxSender();

                worker.DoWork += (workerSender, workerEvent) =>
                {
                    var excelInfo = excelHandler.GetRowsInfo();

                    foreach ((var faxNum, var recipientName) in excelInfo)
                    {
                        if (worker.CancellationPending == true)
                        {
                            workerEvent.Cancel = true;
                            break;
                        }

                        if (documents.TryGetValue(recipientName, out var docPath))
                        {
                            var recipientInfo = new Dictionary<string, string> { { faxNum, recipientName } };
                            faxSender.SendFax(recipientInfo, documents[recipientName]);
                        }

                        worker.ReportProgress((int)(1M / excelInfo.Count() * 100));
                    }
                };

                worker.ProgressChanged += (workerSender, workerEvent) =>
                {
                    ProgressBar.Value = workerEvent.ProgressPercentage;
                };

                worker.RunWorkerCompleted += (workerSender, workerEvent) =>
                {
                    faxSender.DisconnectFaxServer();

                    ProgressBar.Value = 100;

                    System.Windows.Forms.MessageBox.Show($"The current fax process has completed. You may now close the window. :)", "Fax Process", MessageBoxButtons.OK);

                    // Write the log file
                    using (StreamWriter writer = new StreamWriter($"./Log/FaxLog_{DateTime.Now.ToString("yyyyMMddHHmmss")}.txt"))
                    {
                        writer.WriteLine(this.ProcessLogTextBox.Text);
                    }

                    // Enable the Send Fax Button once the worker is terminated or completed
                    //this.SendFaxButton.IsEnabled = true;
                };

                worker.RunWorkerAsync();
            }
        }

        private void IsGenerateWordDoc_CheckBox_Checked(object sender, RoutedEventArgs e)
        {
            Label_DocDir.Content = "Word Template";
            this.selectedFolder = String.Empty;
            this.selectedExcel = String.Empty;
            this.DirTextBox.Text = this.selectedFolder;
            this.ExcelFileTextBox.Text = this.selectedExcel;
            this.DirButton.Content = "Choose Template";

            this.SendFaxButton.Content = "Gen. Word Doc.";

            this.EnableSendFaxButton();
            //isGenerateWordDoc = IsGenerateWordDoc_CheckBox.IsChecked ?? false;
        }

        private void IsGenerateWordDoc_CheckBox_UnChecked(object sender, RoutedEventArgs e)
        {
            Label_DocDir.Content = "Document Folder";
            this.selectedFolder = String.Empty;
            this.selectedExcel = String.Empty;
            this.DirTextBox.Text = this.selectedFolder;
            this.ExcelFileTextBox.Text = this.selectedExcel;
            this.DirButton.Content = "Choose Directory";

            this.SendFaxButton.Content = "Send Fax";

            this.EnableSendFaxButton();
            //isGenerateWordDoc = IsGenerateWordDoc_CheckBox.IsChecked ?? false;
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
            //textbox.Text += value;

            base.Write(value);
            textbox.Dispatcher.BeginInvoke(new Action(() =>
            {
                textbox.AppendText(value.ToString());
            }));
        }

        public override void Write(string value)
        {
            //textbox.Text += value;

            base.Write(value);
            textbox.Dispatcher.BeginInvoke(new Action(() =>
            {
                textbox.AppendText(value.ToString());
            }));
        }

        public override Encoding Encoding
        {
            get { return Encoding.UTF8; }
        }
    }
}
