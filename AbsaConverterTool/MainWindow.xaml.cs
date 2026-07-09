using AbsaConverterTool.Helper;
using Microsoft.WindowsAPICodePack.Dialogs;
using System.Data;
using System.IO;
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
using static System.Windows.Forms.VisualStyles.VisualStyleElement.TextBox;

namespace AbsaConverterTool;


public partial class MainWindow : Window
{
    public MainWindow()
    {
        InitializeComponent();
    }

    private class FileConversionEntry
    {
        public string FilePath { get; set; }
        public DataTable Table { get; set; }
        public List<string> ParseErrors { get; set; } = new();
    }

    private class MdbCreationResult
    {
        public string FilePath { get; set; }
        public string OutputPath { get; set; }
        public int RowCount { get; set; }
        public string Status { get; set; }
        public string Detail { get; set; }
    }

    string sharedPath;
    List<FileConversionEntry> fileEntries = new();

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

        ErrorLabel.Visibility = Visibility.Collapsed;
        LoadingBar.Visibility = Visibility.Visible;
        SuccessLabel.Visibility = Visibility.Collapsed;
        ErrorsListBox.ItemsSource = null;
        ErrorsListBox.Visibility = Visibility.Collapsed;

        var newFileEntries = new List<FileConversionEntry>();
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

                    var (table, errors) = FileHelper.ParseRecordsOfFileToDataTable(file.RecordsInText, file.RecordsInBytes, file.FilePath);

                    newFileEntries.Add(new FileConversionEntry { FilePath = file.FilePath, Table = table, ParseErrors = errors });

                    allErrors.AddRange(errors);
                }
            }
            catch (Exception ex)
            {
                allErrors.Add($"Unexpected error: {ex.Message}");
            }
        });

        fileEntries = newFileEntries;

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

            foreach (var entry in fileEntries)
            {
                LogTextBox.AppendText($"=== {entry.FilePath} ({entry.Table.Rows.Count} rows) ==={Environment.NewLine}");

                foreach (DataRow row in entry.Table.Rows)
                {
                    StringBuilder sb = new StringBuilder();
                    foreach (DataColumn col in entry.Table.Columns)
                    {
                        sb.Append($"{col.ColumnName}: {row[col]} | ");
                    }
                    sb.AppendLine("********************************************************");
                    LogTextBox.AppendText(sb.ToString() + Environment.NewLine);
                }
            }

            SuccessLabel.Visibility = Visibility.Visible;
            int totalRecords = fileEntries.Sum(e => e.Table.Rows.Count);
            SelectedFileNameTextBlock.Text = $"Success! Records loaded: {totalRecords}";
        }


    }
    private async void processButton_Click(object sender, RoutedEventArgs e)
    {
        // Validate that records are loaded
        if (fileEntries == null || fileEntries.Count == 0)
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

        var results = new List<MdbCreationResult>();

        await Task.Run(() =>
        {
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var entry in fileEntries)
            {
                string baseName = System.IO.Path.GetFileNameWithoutExtension(entry.FilePath) + "_Output.mdb";
                string outputName = baseName;
                int suffix = 1;
                while (!usedNames.Add(outputName))
                {
                    suffix++;
                    outputName = System.IO.Path.GetFileNameWithoutExtension(entry.FilePath) + $"_Output_{suffix}.mdb";
                }

                string outputPath = System.IO.Path.Combine(sharedPath, outputName);

                if (entry.Table == null || entry.Table.Rows.Count == 0)
                {
                    results.Add(new MdbCreationResult { FilePath = entry.FilePath, Status = "Skipped", Detail = "0 valid records parsed" });
                    continue;
                }

                try
                {
                    FileHelper.CreateMdbFromDataTable(entry.Table, outputPath);
                    results.Add(new MdbCreationResult { FilePath = entry.FilePath, OutputPath = outputPath, RowCount = entry.Table.Rows.Count, Status = "Created" });
                }
                catch (Exception ex)
                {
                    results.Add(new MdbCreationResult { FilePath = entry.FilePath, OutputPath = outputPath, Status = "Failed", Detail = ex.Message });
                }
            }
        });

        LoadingBar.Visibility = Visibility.Collapsed;

        LogTextBox.Clear();
        foreach (var result in results)
        {
            string line = result.Status switch
            {
                "Created" => $"{result.FilePath} -> Created: {result.OutputPath} ({result.RowCount} records)",
                "Skipped" => $"{result.FilePath} -> Skipped: {result.Detail}",
                _ => $"{result.FilePath} -> Failed: {result.Detail}",
            };
            LogTextBox.AppendText(line + Environment.NewLine);
        }

        int createdCount = results.Count(r => r.Status == "Created");
        int skippedCount = results.Count(r => r.Status == "Skipped");
        int failedCount = results.Count(r => r.Status == "Failed");

        if (skippedCount + failedCount > 0)
        {
            ErrorLabel.Content = $"{failedCount} failed, {skippedCount} skipped out of {results.Count} files";
            ErrorLabel.Visibility = Visibility.Visible;

            ErrorsListBox.ItemsSource = results
                .Where(r => r.Status != "Created")
                .Select(r => $"File: {r.FilePath}\n{r.Status}: {r.Detail}")
                .ToList();
            ErrorsListBox.Visibility = Visibility.Visible;
            FailedP3Panel.Visibility = Visibility.Visible;
        }

        if (createdCount > 0)
        {
            SuccessLabel.Visibility = Visibility.Visible;
        }

        SelectedFileNameTextBlock.Text = $"Created {createdCount} of {results.Count} mdb file(s) in {sharedPath}";
    }

}
