using AbsaConverterTool.Helper;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Data;
using System.IO;
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;

namespace AbsaConverterTool;


public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }
    public List<byte[]> FilesRecordContentInBytes { get; set; } = new List<byte[]>();
    public List<string> FilesRecordContentInText { get; set; } = new List<string>();
    string sharedPath, errorMsg = string.Empty;
    DataTable cards;

    #region Code-Behind for Moving Window
    private void TitleBar_MouseDown(object sender, MouseButtonEventArgs e)
    {
        if (e.ChangedButton == MouseButton.Left)
            this.DragMove();
    }

    private void Minimize_Click(object sender, RoutedEventArgs e)
    {
        this.WindowState = WindowState.Minimized;
    }

    private void Maximize_Click(object sender, RoutedEventArgs e)
    {
        this.WindowState = this.WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
    }

    private void Close_Click(object sender, RoutedEventArgs e)
    {
        this.Close();
    }

    #endregion
    
    private async void ChooseFolder_Click(object sender, RoutedEventArgs e)
    {
        var dialog = new System.Windows.Forms.FolderBrowserDialog();
        dialog.Description = "Select a folder containing .inp files";

        var result = dialog.ShowDialog();
        if (result != System.Windows.Forms.DialogResult.OK) return;

        string folderPath = dialog.SelectedPath;
        this.sharedPath = folderPath;

        FilesRecordContentInBytes.Clear();
        FilesRecordContentInText.Clear();
        ErrorLabel.Visibility = Visibility.Collapsed;
        LoadingBar.Visibility = Visibility.Visible;
        SuccessLabel.Visibility = Visibility.Collapsed;
        ErrorsListBox.ItemsSource = null;
        ErrorsListBox.Visibility = Visibility.Collapsed;

        cards = new DataTable("Cards");
        var allErrors = new List<string>();

        await Task.Run(() =>
        {
            try
            {
                var fileResults = FileHelper.GetAllFilesRecords(folderPath); // returns list of {FilePath, RecordsInBytes, RecordsInText, ErrorMessage}

                foreach (var file in fileResults)
                {
                    if (!string.IsNullOrEmpty(file.ErrorMessage))
                    {
                        allErrors.Add($"File: {file.FilePath}\nError: {file.ErrorMessage}");
                        continue;
                    }

                    FilesRecordContentInBytes.AddRange(file.RecordsInBytes);
                    FilesRecordContentInText.AddRange(file.RecordsInText);

                    var (table, errors) = FileHelper.ParseRecordsOfFileToDataTable(file.RecordsInText, file.RecordsInBytes, file.FilePath);

                    foreach (DataColumn col in table.Columns)
                    {
                        if (!cards.Columns.Contains(col.ColumnName))
                            cards.Columns.Add(col.ColumnName, col.DataType);
                    }

                    foreach (DataRow row in table.Rows)
                        cards.ImportRow(row);

                    allErrors.AddRange(errors);
                }
            }
            catch (Exception ex)
            {
                allErrors.Add($"Unexpected error: {ex.Message}");
            }
        });

        LoadingBar.Visibility = Visibility.Collapsed;

        if (allErrors.Count > 0)
        {
            ErrorLabel.Content = $"Errors found: {allErrors.Count} file(s)/record(s)";
            ErrorLabel.Visibility = Visibility.Visible;

            ErrorsListBox.ItemsSource = allErrors;
            ErrorsListBox.Visibility = Visibility.Visible;
            FailedP3Panel.Visibility = Visibility.Visible;
        }
        else
        {
            LogTextBox.Clear(); // clear previous logs

            foreach (DataRow row in cards.Rows)
            {
                StringBuilder sb = new StringBuilder();
                foreach (DataColumn col in cards.Columns)
                {
                    sb.Append($"{col.ColumnName}: {row[col]} | ");
                }
                sb.AppendLine("********************************************************");
                // Append line to TextBox on UI thread
                Dispatcher.Invoke(() =>
                {
                    LogTextBox.AppendText(sb.ToString() + Environment.NewLine);
                });
            }

            SuccessLabel.Visibility = Visibility.Visible;
            SelectedFileNameTextBlock.Text = $"Success! Records loaded: {FilesRecordContentInText.Count}";
        }


    }
    private async void processButton_Click(object sender, RoutedEventArgs e)
    {
        // Validate that records are loaded
        if (FilesRecordContentInText == null || FilesRecordContentInText.Count == 0)
        {
            ErrorLabel.Content = "No records loaded. Please choose a folder first.";
            ErrorLabel.Visibility = Visibility.Visible;
            return;
        }

        // Clear UI and show loading
        LogTextBox.Clear();
        ErrorLabel.Visibility = Visibility.Collapsed;
        SuccessLabel.Visibility = Visibility.Collapsed;
        LoadingBar.Visibility = Visibility.Visible;

        string errorMsg = string.Empty;

        await Task.Run(() =>
        {
            try
            {
                // Convert loaded records to DataTable
               

                // MDB path
                string mdbPath = System.IO.Path.Combine(sharedPath, "CardsOutput.mdb");

               

                // Create MDB file from DataTable
                FileHelper.CreateMdbFromDataTable(cards, mdbPath);
            }
            catch (Exception ex)
            {
                errorMsg = ex.Message;
            }
        });

        LoadingBar.Visibility = Visibility.Collapsed;

        if (!string.IsNullOrEmpty(errorMsg))
        {
            ErrorLabel.Content = errorMsg;
            ErrorLabel.Visibility = Visibility.Visible;
            LogTextBox.Text = errorMsg;
        }
        else
        {
            SuccessLabel.Visibility = Visibility.Visible;
            SelectedFileNameTextBlock.Text = $"Success! MDB created at: {System.IO.Path.Combine(sharedPath, "CardsOutput.mdb")}";

            // Show DataTable records in LogTextBox
            LogTextBox.Clear();
            foreach (DataRow row in cards.Rows)
            {
                StringBuilder sb = new StringBuilder();
                foreach (DataColumn col in cards.Columns)
                {
                    sb.Append($"{col.ColumnName}: {row[col]} | ");
                }
                LogTextBox.AppendText(sb.ToString() + Environment.NewLine);
            }
        }
    }

}
