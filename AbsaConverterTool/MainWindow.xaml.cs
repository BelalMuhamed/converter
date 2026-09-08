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
        public FileParseResult ParseResult { get; set; }
    }

    private class MdbCreationResult
    {
        public string FilePath { get; set; }
        public string OutputPath { get; set; }
        public string ProductDisplayName { get; set; }
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

                    var parseResult = FileHelper.ParseFileRecords(file.RecordsInText, file.RecordsInBytes, file.FilePath);

                    newFileEntries.Add(new FileConversionEntry { FilePath = file.FilePath, ParseResult = parseResult });

                    // Only hard parse exceptions surface here, same as before. Records that fail
                    // because of an unknown/missing product code or oversized chip data are not
                    // exceptions - they are carried in ParseResult.FailedRecords for the Excel
                    // report generated during "Process Files".
                    allErrors.AddRange(parseResult.ParseErrors);
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
                int failedCount = entry.ParseResult.FailedRecords.Count;
                LogTextBox.AppendText($"=== {entry.FilePath} ({entry.ParseResult.Records.Count} valid record(s), {failedCount} failed record(s)) ==={Environment.NewLine}");

                foreach (var record in entry.ParseResult.Records)
                {
                    ProductMapping.TryGet(record.ProductCode, out var productInfo);
                    StringBuilder sb = new StringBuilder();
                    sb.Append($"Product: {productInfo?.DisplayName ?? record.ProductCode} ({record.ProductCode}) | ");
                    sb.Append($"IDWAutoNumber: {record.RecordIndex} | ");
                    sb.Append($"IDWPAN: {record.PAN} | ");
                    sb.Append($"IDWEXP: {record.EXP} | ");
                    sb.Append($"IDWNAME: {record.Name} | ");
                    sb.Append($"IDWCVV2: {record.CVV2} | ");
                    sb.Append($"IDWTrack1: {record.Track1} | ");
                    sb.Append($"IDWTrack2: {record.Track2} | ");
                    sb.Append($"IDWChip1: {record.Chip1} | ");
                    sb.Append($"IDWChip2: {record.Chip2} | ");
                    sb.Append($"IDWChip3: {record.Chip3} | ");
                    sb.Append($"IDWChip4: {record.Chip4} | ");
                    sb.AppendLine("********************************************************");
                    LogTextBox.AppendText(sb.ToString() + Environment.NewLine);
                }
            }

            SuccessLabel.Visibility = Visibility.Visible;
            int totalRecords = fileEntries.Sum(e => e.ParseResult.Records.Count);
            int totalFailed = fileEntries.Sum(e => e.ParseResult.FailedRecords.Count);
            SelectedFileNameTextBlock.Text = totalFailed > 0
                ? $"Success! Valid records loaded: {totalRecords} ({totalFailed} will be reported as failed in Failures.xlsx)"
                : $"Success! Records loaded: {totalRecords}";
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
        var allFailedRecords = new List<FailedRecord>();
        string failuresReportPath = null;

        await Task.Run(() =>
        {
            var usedNames = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (var entry in fileEntries)
            {
                allFailedRecords.AddRange(entry.ParseResult.FailedRecords);

                // One output .mdb per Product Code bucket found in this file. AGVQ (Signature
                // Credit + Signature Debit) is a single bucket by construction, since both
                // share the same ProductCode.
                var groups = entry.ParseResult.Records
                    .GroupBy(r => r.ProductCode)
                    .ToList();

                if (groups.Count == 0)
                {
                    results.Add(new MdbCreationResult { FilePath = entry.FilePath, Status = "Skipped", Detail = "0 valid records parsed" });
                    continue;
                }

                string baseName = System.IO.Path.GetFileNameWithoutExtension(entry.FilePath);

                foreach (var group in groups)
                {
                    ProductMapping.TryGet(group.Key, out var productInfo);
                    string token = productInfo?.FileNameToken ?? group.Key;

                    string outputName = $"{baseName}_{token}.mdb";
                    int suffix = 1;
                    while (!usedNames.Add(outputName))
                    {
                        suffix++;
                        outputName = $"{baseName}_{token}_{suffix}.mdb";
                    }

                    string outputPath = System.IO.Path.Combine(sharedPath, outputName);
                    var recordsForProduct = group.ToList();

                    try
                    {
                        var table = FileHelper.BuildCardsDataTable(recordsForProduct);
                        FileHelper.CreateMdbFromDataTable(table, outputPath);
                        results.Add(new MdbCreationResult
                        {
                            FilePath = entry.FilePath,
                            OutputPath = outputPath,
                            ProductDisplayName = productInfo?.DisplayName ?? group.Key,
                            RowCount = recordsForProduct.Count,
                            Status = "Created"
                        });
                    }
                    catch (Exception ex)
                    {
                        results.Add(new MdbCreationResult
                        {
                            FilePath = entry.FilePath,
                            OutputPath = outputPath,
                            ProductDisplayName = productInfo?.DisplayName ?? group.Key,
                            Status = "Failed",
                            Detail = ex.Message
                        });
                    }
                }
            }

            if (allFailedRecords.Count > 0)
            {
                failuresReportPath = System.IO.Path.Combine(sharedPath, "Failures.xlsx");
                FileHelper.WriteFailuresReport(allFailedRecords, failuresReportPath);
            }
        });

        LoadingBar.Visibility = Visibility.Collapsed;

        LogTextBox.Clear();
        foreach (var result in results)
        {
            string line = result.Status switch
            {
                "Created" => $"{result.FilePath} -> Created: {result.OutputPath} [{result.ProductDisplayName}] ({result.RowCount} records)",
                "Skipped" => $"{result.FilePath} -> Skipped: {result.Detail}",
                _ => $"{result.FilePath} -> Failed: {result.Detail}",
            };
            LogTextBox.AppendText(line + Environment.NewLine);
        }

        if (allFailedRecords.Count > 0)
        {
            LogTextBox.AppendText(Environment.NewLine + $"{allFailedRecords.Count} record(s) failed validation -> see {failuresReportPath}" + Environment.NewLine);
        }

        int createdCount = results.Count(r => r.Status == "Created");
        int skippedCount = results.Count(r => r.Status == "Skipped");
        int failedCount = results.Count(r => r.Status == "Failed");

        if (skippedCount + failedCount > 0)
        {
            ErrorLabel.Content = $"{failedCount} failed, {skippedCount} skipped out of {results.Count} output file(s)";
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

        string failuresNote = allFailedRecords.Count > 0
            ? $", {allFailedRecords.Count} record(s) failed (see Failures.xlsx)"
            : string.Empty;

        SelectedFileNameTextBlock.Text = $"Created {createdCount} of {results.Count} mdb file(s) in {sharedPath}{failuresNote}";
    }

}
