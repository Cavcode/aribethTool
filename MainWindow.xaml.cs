using System;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Threading.Tasks;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Input;
using Microsoft.Win32;

namespace aribeth
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string? _currentFilePath;
        private DataTable? _twoDaTable;
        private readonly List<List<string>> _rowClipboard = new();
        private int _twoDaIndexWidth;
        private int _twoDaHeaderIndent = 1;
        private List<int> _twoDaColumnWidths = new();
        private List<int> _twoDaHeaderTokenStarts = new();
        private int _twoDaHeaderVisualLength;
        private readonly Dictionary<string, (List<int> Starts, int VisualLength)> _twoDaRowLayouts = new();
        private string? _nwn2daPath;
        private string? _currentTlkPath;
        private string? _nwnTlkPath;
        private List<TlkEntry> _tlkEntries = new();
        private string _tlkJsonCache = string.Empty;
        private const int TlkBaseIndex = 1677216;
        private readonly Stack<string> _twoDaUndo = new();
        private readonly Stack<string> _twoDaRedo = new();
        private readonly Stack<string> _tlkUndo = new();
        private readonly Stack<string> _tlkRedo = new();
        private bool _suppressUndoCapture;
        private bool _isTlkLoading;
        private const int TlkUserIndexBase = 16777216;
        private JsonObject? _tlkJsonRoot;

        public MainWindow()
        {
            InitializeComponent();
            DataGrid2DA.IsReadOnly = false;
            DataGridTLK.ItemsSource = _tlkEntries;
            DataGridTLK.CellEditEnding += DataGridTLK_CellEditEnding;
            UpdateEmptyStateImages();
        }

        private void UpdateEmptyStateImages()
        {
            if (TwoDaEmptyStateImage != null)
            {
                TwoDaEmptyStateImage.Visibility = _twoDaTable == null ? Visibility.Visible : Visibility.Collapsed;
            }

            if (TlkEmptyStateImage != null)
            {
                TlkEmptyStateImage.Visibility = string.IsNullOrWhiteSpace(_currentTlkPath)
                    ? Visibility.Visible
                    : Visibility.Collapsed;
            }
        }

        private void SetStatus(string message)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => StatusTextBlock.Text = message);
                return;
            }

            StatusTextBlock.Text = message;
        }

        private void Open2DAMenuItem_Click(object sender, RoutedEventArgs e) => Open2DAFile();
        private void OpenTLKMenuItem_Click(object sender, RoutedEventArgs e) => OpenTlkFile();
        private void CloseCurrentMenuItem_Click(object sender, RoutedEventArgs e) => CloseCurrent();
        private void SaveMenuItem_Click(object sender, RoutedEventArgs e) => SaveCurrent();
        private void SaveAsMenuItem_Click(object sender, RoutedEventArgs e) => SaveCurrent(saveAs: true);
        private void UndoMenuItem_Click(object sender, RoutedEventArgs e) => Undo();
        private void RedoMenuItem_Click(object sender, RoutedEventArgs e) => Redo();
        private void ExitMenuItem_Click(object sender, RoutedEventArgs e) => Close();
        private void CopyRowMenuItem_Click(object sender, RoutedEventArgs e) => CopyRows(cut: false);
        private void CutRowMenuItem_Click(object sender, RoutedEventArgs e) => CopyRows(cut: true);
        private void PasteRowMenuItem_Click(object sender, RoutedEventArgs e) => PasteRows();
        private void AddColumnMenuItem_Click(object sender, RoutedEventArgs e) => AddColumn();
        private void RemoveColumnMenuItem_Click(object sender, RoutedEventArgs e) => RemoveColumn();
        private void FindMenuItem_Click(object sender, RoutedEventArgs e)
        {
            SearchTextBox.Focus();
            SearchTextBox.SelectAll();
            SetStatus("Find");
        }
        private void ReplaceMenuItem_Click(object sender, RoutedEventArgs e)
        {
            ReplaceTextBox.Focus();
            ReplaceTextBox.SelectAll();
        }
        private void Normalize2DAMenuItem_Click(object sender, RoutedEventArgs e) => Normalize2DA();
        private void Merge2DAMenuItem_Click(object sender, RoutedEventArgs e) => Merge2DA();
        private void CheckIssuesMenuItem_Click(object sender, RoutedEventArgs e) => Check2DAIssues();
        private void ExportTLKToJsonMenuItem_Click(object sender, RoutedEventArgs e) => ExportTlkToJson();
        private void ImportTLKFromJsonMenuItem_Click(object sender, RoutedEventArgs e) => ImportTlkFromJson();
        private void ExportTLKToBinaryMenuItem_Click(object sender, RoutedEventArgs e) => ExportTlkToBinary();

        private void OpenButton_Click(object sender, RoutedEventArgs e) => OpenCommand_Executed(sender, null!);
        private void SaveButton_Click(object sender, RoutedEventArgs e) => SaveCurrent();
        private void CloseCurrentButton_Click(object sender, RoutedEventArgs e) => CloseCurrent();
        private void CopyButton_Click(object sender, RoutedEventArgs e) => CopyRows(cut: false);
        private void PasteButton_Click(object sender, RoutedEventArgs e) => PasteRows();
        private void CutButton_Click(object sender, RoutedEventArgs e) => CopyRows(cut: true);
        private void FindButton_Click(object sender, RoutedEventArgs e)
        {
            SearchTextBox.Focus();
            SearchTextBox.SelectAll();
            SetStatus("Find");
        }
        private void ReplaceButton_Click(object sender, RoutedEventArgs e)
        {
            ReplaceTextBox.Focus();
            ReplaceTextBox.SelectAll();
            ReplaceValues();
        }
        private void AddColumnButton_Click(object sender, RoutedEventArgs e) => AddColumn();
        private void RemoveColumnButton_Click(object sender, RoutedEventArgs e) => RemoveColumn();

        private void ColumnFilterTextBox_TextChanged(object sender, TextChangedEventArgs e) => ApplyColumnFilter();
        private void SearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            ApplyRowFilter();
            SelectFirstMatchIn2da();
        }
        private void SearchTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Right)
            {
                FindNextMatchIn2da(forward: true);
                e.Handled = true;
            }
            else if (e.Key == Key.Left)
            {
                FindNextMatchIn2da(forward: false);
                e.Handled = true;
            }
        }
        private void ClearFiltersButton_Click(object sender, RoutedEventArgs e) => ClearFilters();

        private void CopyRowContextMenuItem_Click(object sender, RoutedEventArgs e) => CopyRows(cut: false);
        private void CutRowContextMenuItem_Click(object sender, RoutedEventArgs e) => CopyRows(cut: true);
        private void PasteRowContextMenuItem_Click(object sender, RoutedEventArgs e) => PasteRows();
        private void InsertRowAboveContextMenuItem_Click(object sender, RoutedEventArgs e) => InsertRow(above: true);
        private void InsertRowBelowContextMenuItem_Click(object sender, RoutedEventArgs e) => InsertRow(above: false);
        private void DeleteRowContextMenuItem_Click(object sender, RoutedEventArgs e) => DeleteSelectedRows();

        private void TLKSearchTextBox_TextChanged(object sender, TextChangedEventArgs e)
        {
            SelectFirstMatchInTlk();
            SetStatus("TLK search updated");
        }
        private void TLKSearchTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Right)
            {
                FindNextMatchInTlk(forward: true);
                e.Handled = true;
            }
            else if (e.Key == Key.Left)
            {
                FindNextMatchInTlk(forward: false);
                e.Handled = true;
            }
        }
        private void TLKIndexTextBox_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter)
            {
                GoToTlkIndex();
                e.Handled = true;
            }
        }
        private void TLKClearFiltersButton_Click(object sender, RoutedEventArgs e) => SetStatus("TLK filters cleared");
        private void TLKRawModeCheckBox_Checked(object sender, RoutedEventArgs e) => SetTlkRawMode(true);
        private void TLKRawModeCheckBox_Unchecked(object sender, RoutedEventArgs e) => SetTlkRawMode(false);
        private void TLKGoToIndexButton_Click(object sender, RoutedEventArgs e) => GoToTlkIndex();

        private void Command_CanExecute(object sender, CanExecuteRoutedEventArgs e) => e.CanExecute = true;
        private void UndoCommand_Executed(object sender, ExecutedRoutedEventArgs e) => Undo();
        private void RedoCommand_Executed(object sender, ExecutedRoutedEventArgs e) => Redo();
        private void CopyCommand_Executed(object sender, ExecutedRoutedEventArgs e) => CopyRows(cut: false);
        private void CutCommand_Executed(object sender, ExecutedRoutedEventArgs e) => CopyRows(cut: true);
        private void PasteCommand_Executed(object sender, ExecutedRoutedEventArgs e) => PasteRows();
        private void FindCommand_Executed(object sender, ExecutedRoutedEventArgs e) => FindMenuItem_Click(sender, e);
        private void ReplaceCommand_Executed(object sender, ExecutedRoutedEventArgs e) => ReplaceMenuItem_Click(sender, e);
        private void OpenCommand_Executed(object sender, ExecutedRoutedEventArgs e)
        {
            var dialog = new OpenFileDialog
            {
                Filter = "2DA or TLK files (*.2da;*.tlk)|*.2da;*.tlk|2DA files (*.2da)|*.2da|TLK files (*.tlk)|*.tlk|All files (*.*)|*.*",
                Title = "Open 2DA or TLK File"
            };

            if (dialog.ShowDialog() != true)
            {
                return;
            }

            var extension = Path.GetExtension(dialog.FileName);
            if (extension.Equals(".tlk", StringComparison.OrdinalIgnoreCase))
            {
                _ = OpenTlkFileAsync(dialog.FileName);
            }
            else
            {
                Open2DAFile(dialog.FileName);
            }
        }

        private void SaveCommand_Executed(object sender, ExecutedRoutedEventArgs e) => SaveCurrent();
        private void SaveAsCommand_Executed(object sender, ExecutedRoutedEventArgs e) => SaveCurrent(saveAs: true);

        private void TitleBar_MouseLeftButtonDown(object sender, MouseButtonEventArgs e)
        {
            if (e.ClickCount == 2)
            {
                WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
                return;
            }

            DragMove();
        }

        private void MinimizeButton_Click(object sender, RoutedEventArgs e) => WindowState = WindowState.Minimized;
        private void MaximizeButton_Click(object sender, RoutedEventArgs e) =>
            WindowState = WindowState == WindowState.Maximized ? WindowState.Normal : WindowState.Maximized;
        private void CloseButton_Click(object sender, RoutedEventArgs e) => Close();

        private void SaveCurrent(bool saveAs = false)
        {
            if (IsTlkTabActive())
            {
                SaveTlkFile(saveAs);
            }
            else
            {
                Save2DAFile(saveAs);
            }
        }

        private bool IsTlkTabActive()
        {
            if (MainTabControl.SelectedItem is TabItem item)
            {
                return string.Equals(item.Header?.ToString(), "TLK Editor", StringComparison.OrdinalIgnoreCase);
            }

            return false;
        }

        private void CloseCurrent()
        {
            if (IsTlkTabActive())
            {
                _currentTlkPath = null;
                _tlkEntries = new List<TlkEntry>();
                DataGridTLK.ItemsSource = _tlkEntries;
                TLKRawTextBox.Text = string.Empty;
                _tlkJsonCache = string.Empty;
                _tlkJsonRoot = null;
                SetStatus("Closed TLK");
            }
            else
            {
                _currentFilePath = null;
                _twoDaTable = null;
                _twoDaIndexWidth = 0;
                _twoDaHeaderIndent = 1;
                _twoDaColumnWidths = new List<int>();
                _twoDaHeaderTokenStarts = new List<int>();
                _twoDaHeaderVisualLength = 0;
                _twoDaRowLayouts.Clear();
                DataGrid2DA.ItemsSource = null;
                SetStatus("Closed 2DA");
            }

            FileInfoTextBlock.Text = "No file loaded";
            UpdateEmptyStateImages();
        }

        private static bool FileLooksLike2da(string filePath)
        {
            try
            {
                using var reader = new StreamReader(filePath);
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    if (line == null)
                    {
                        break;
                    }

                    if (string.IsNullOrWhiteSpace(line))
                    {
                        continue;
                    }

                    return line.TrimStart().StartsWith("2DA", StringComparison.OrdinalIgnoreCase);
                }
            }
            catch
            {
                return false;
            }

            return false;
        }

        private void Open2DAFile(string? filePath = null)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                var dialog = new OpenFileDialog
                {
                    Filter = "2DA files (*.2da)|*.2da|All files (*.*)|*.*",
                    Title = "Open 2DA File"
                };

                if (dialog.ShowDialog() != true)
                {
                    return;
                }

                filePath = dialog.FileName;
            }

            if (filePath.EndsWith(".tlk", StringComparison.OrdinalIgnoreCase))
            {
                _ = OpenTlkFileAsync(filePath);
                return;
            }

            try
            {
                if (!FileLooksLike2da(filePath))
                {
                    MessageBox.Show("This file does not appear to be a valid 2DA file.", "Aribeth",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                var parsed = TwoDAParser.Parse(filePath, LogMessage, strictHeader: true);
                _twoDaTable = ToDataTable(parsed);
                UpdateTwoDaFormattingState(parsed);
                DataGrid2DA.ItemsSource = _twoDaTable.DefaultView;
                HookTwoDaEvents();
                CaptureTwoDaSnapshot(clearRedo: true);
                _currentFilePath = filePath;
                FileInfoTextBlock.Text = filePath;
                SetStatus($"Loaded 2DA: {Path.GetFileName(filePath)}");
                UpdateEmptyStateImages();
            }
            catch (Exception ex)
            {
                LogMessage($"Failed to open 2DA: {ex.Message}");
                SetStatus("Failed to open 2DA");
            }
        }

        private void Save2DAFile(bool saveAs = false)
        {
            if (_twoDaTable == null)
            {
                SetStatus("No 2DA loaded");
                return;
            }

            if (saveAs || string.IsNullOrWhiteSpace(_currentFilePath))
            {
                var saveDialog = new SaveFileDialog
                {
                    Filter = "2DA files (*.2da)|*.2da|All files (*.*)|*.*",
                    Title = "Save 2DA File"
                };

                if (saveDialog.ShowDialog() != true)
                {
                    return;
                }

                _currentFilePath = saveDialog.FileName;
            }

            try
            {
                var data = BuildTwoDaDataFromTable();
                var contents = TwoDAParser.Serialize(data);
                File.WriteAllText(_currentFilePath!, contents);
                FileInfoTextBlock.Text = _currentFilePath!;
                SetStatus($"Saved 2DA: {Path.GetFileName(_currentFilePath!)}");
            }
            catch (Exception ex)
            {
                LogMessage($"Failed to save 2DA: {ex.Message}");
                SetStatus("Failed to save 2DA");
            }
        }

        private static DataTable ToDataTable(TwoDAData data)
        {
            var table = new DataTable();
            table.Columns.Add("Row", typeof(string));

            foreach (var column in data.Columns)
            {
                table.Columns.Add(column, typeof(string));
            }

            foreach (var row in data.Rows)
            {
                var dataRow = table.NewRow();
                dataRow["Row"] = row.Index;
                for (var i = 0; i < data.Columns.Count; i++)
                {
                    dataRow[data.Columns[i]] = row.Values[i];
                }

                table.Rows.Add(dataRow);
            }

            return table;
        }

        private static TwoDAData FromDataTable(DataTable table)
        {
            var data = new TwoDAData();
            foreach (DataColumn column in table.Columns)
            {
                if (column.ColumnName == "Row")
                {
                    continue;
                }

                data.Columns.Add(column.ColumnName);
            }

            foreach (DataRow row in table.Rows)
            {
                var newRow = new TwoDARow
                {
                    Index = row["Row"]?.ToString() ?? string.Empty
                };

                foreach (var column in data.Columns)
                {
                    var value = row[column]?.ToString();
                    newRow.Values.Add(string.IsNullOrWhiteSpace(value) ? "****" : value);
                }

                data.Rows.Add(newRow);
            }

            return data;
        }

        private TwoDAData BuildTwoDaDataFromTable()
        {
            var data = FromDataTable(_twoDaTable!);
            ApplyTwoDaFormatting(data);
            return data;
        }

        private void ApplyTwoDaFormatting(TwoDAData data)
        {
            data.IndexWidth = _twoDaIndexWidth;
            data.HeaderIndent = _twoDaHeaderIndent;
            data.ColumnWidths.Clear();
            data.ColumnWidths.AddRange(_twoDaColumnWidths);
            data.HeaderTokenStarts.Clear();
            data.HeaderTokenStarts.AddRange(_twoDaHeaderTokenStarts);
            data.HeaderVisualLength = _twoDaHeaderVisualLength;

            foreach (var row in data.Rows)
            {
                row.TokenStarts.Clear();
                row.VisualLength = 0;
                if (!string.IsNullOrWhiteSpace(row.Index) && _twoDaRowLayouts.TryGetValue(row.Index, out var layout))
                {
                    row.TokenStarts.AddRange(layout.Starts);
                    row.VisualLength = layout.VisualLength;
                }
            }
        }

        private void UpdateTwoDaFormattingState(TwoDAData data)
        {
            _twoDaIndexWidth = data.IndexWidth;
            _twoDaHeaderIndent = data.HeaderIndent;
            _twoDaColumnWidths = data.ColumnWidths.ToList();
            _twoDaHeaderTokenStarts = data.HeaderTokenStarts.ToList();
            _twoDaHeaderVisualLength = data.HeaderVisualLength;
            _twoDaRowLayouts.Clear();
            foreach (var row in data.Rows)
            {
                if (!string.IsNullOrWhiteSpace(row.Index) && row.TokenStarts.Count > 0)
                {
                    _twoDaRowLayouts[row.Index] = (row.TokenStarts.ToList(), row.VisualLength);
                }
            }
        }

        private void LogMessage(string message)
        {
            if (!Dispatcher.CheckAccess())
            {
                Dispatcher.Invoke(() => LogMessage(message));
                return;
            }

            LogTextBox.AppendText($"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
            LogTextBox.ScrollToEnd();
        }

        private void SetTlkRawMode(bool enabled)
        {
            TLKRawTextBox.Visibility = enabled ? Visibility.Visible : Visibility.Collapsed;
            DataGridTLK.Visibility = enabled ? Visibility.Collapsed : Visibility.Visible;
        }

        private void OpenTlkFile() => _ = OpenTlkFileAsync();

        private async Task OpenTlkFileAsync(string? filePath = null)
        {
            if (_isTlkLoading)
            {
                SetStatus("TLK load already in progress");
                return;
            }

            if (string.IsNullOrWhiteSpace(filePath))
            {
                var dialog = new OpenFileDialog
                {
                    Filter = "TLK files (*.tlk)|*.tlk|All files (*.*)|*.*",
                    Title = "Open TLK File"
                };

                if (dialog.ShowDialog() != true)
                {
                    return;
                }

                filePath = dialog.FileName;
            }

            if (filePath.EndsWith(".2da", StringComparison.OrdinalIgnoreCase))
            {
                Open2DAFile(filePath);
                return;
            }

            var toolPath = ResolveNwnTlkPath();
            if (toolPath == null)
            {
                return;
            }

            try
            {
                _isTlkLoading = true;
                SetStatus("Loading TLK...");

                var tempJson = Path.Combine(Path.GetTempPath(), $"{Path.GetFileNameWithoutExtension(filePath)}.tlk.json");
                var success = await Task.Run(() => RunNwnTlk(
                    toolPath,
                    $"-i \"{filePath}\" -l tlk -o \"{tempJson}\" -k json --quiet",
                    "export json"));

                if (!success)
                {
                    SetStatus("Failed to export TLK json");
                    return;
                }

                var rawJson = await Task.Run(() => File.Exists(tempJson) ? File.ReadAllText(tempJson) : string.Empty);
                var entries = await Task.Run(() => ParseTlkEntriesFromJson(rawJson));

                _currentTlkPath = filePath;
                TLKRawTextBox.Text = rawJson;
                LoadTlkEntries(entries, rawJson);
                TLKRawModeCheckBox.IsChecked = false;
                CaptureTlkSnapshot(clearRedo: true);
                MainTabControl.SelectedItem = MainTabControl.Items
                    .OfType<TabItem>()
                    .FirstOrDefault(tab => string.Equals(tab.Header?.ToString(), "TLK Editor", StringComparison.OrdinalIgnoreCase));
                FileInfoTextBlock.Text = filePath;
                SetStatus($"Loaded TLK: {Path.GetFileName(filePath)}");
                UpdateEmptyStateImages();
            }
            finally
            {
                _isTlkLoading = false;
            }
        }

        private void ExportTlkToJson()
        {
            var toolPath = ResolveNwnTlkPath();
            if (toolPath == null)
            {
                return;
            }

            var inputPath = _currentTlkPath;
            if (string.IsNullOrWhiteSpace(inputPath))
            {
                var openDialog = new OpenFileDialog
                {
                    Filter = "TLK files (*.tlk)|*.tlk|All files (*.*)|*.*",
                    Title = "Select TLK File"
                };

                if (openDialog.ShowDialog() != true)
                {
                    return;
                }

                inputPath = openDialog.FileName;
            }

            var saveDialog = new SaveFileDialog
            {
                Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                Title = "Save TLK JSON",
                FileName = Path.GetFileNameWithoutExtension(inputPath) + ".json"
            };

            if (saveDialog.ShowDialog() != true)
            {
                return;
            }

            var outputPath = saveDialog.FileName;
            var success = RunNwnTlk(
                toolPath,
                $"-i \"{inputPath}\" -l tlk -o \"{outputPath}\" -k json --pretty",
                "export json");

            if (success)
            {
                SetStatus("TLK json exported");
                LogMessage($"Exported TLK to json: {outputPath}");
            }
        }

        private void ImportTlkFromJson()
        {
            var toolPath = ResolveNwnTlkPath();
            if (toolPath == null)
            {
                return;
            }

            var openDialog = new OpenFileDialog
            {
                Filter = "JSON files (*.json)|*.json|All files (*.*)|*.*",
                Title = "Select TLK JSON File"
            };

            if (openDialog.ShowDialog() != true)
            {
                return;
            }

            var inputPath = openDialog.FileName;
            var saveDialog = new SaveFileDialog
            {
                Filter = "TLK files (*.tlk)|*.tlk|All files (*.*)|*.*",
                Title = "Save TLK File",
                FileName = Path.GetFileNameWithoutExtension(inputPath) + ".tlk"
            };

            if (saveDialog.ShowDialog() != true)
            {
                return;
            }

            var outputPath = saveDialog.FileName;
            var success = RunNwnTlk(
                toolPath,
                $"-i \"{inputPath}\" -l json -o \"{outputPath}\" -k tlk --quiet",
                "import json");

            if (success)
            {
                _currentTlkPath = outputPath;
                FileInfoTextBlock.Text = outputPath;
                SetStatus("TLK imported");
                LogMessage($"Imported TLK from text: {outputPath}");
            }
        }

        private void SaveTlkFile(bool saveAs)
        {
            var toolPath = ResolveNwnTlkPath();
            if (toolPath == null)
            {
                return;
            }

            var outputPath = _currentTlkPath;
            if (saveAs || string.IsNullOrWhiteSpace(outputPath))
            {
                var saveDialog = new SaveFileDialog
                {
                    Filter = "TLK files (*.tlk)|*.tlk|All files (*.*)|*.*",
                    Title = "Save TLK File",
                    FileName = "dialog.tlk"
                };

                if (saveDialog.ShowDialog() != true)
                {
                    return;
                }

                outputPath = saveDialog.FileName;
            }

            var baseName = Path.GetFileNameWithoutExtension(outputPath);
            var tempJson = Path.Combine(Path.GetTempPath(), $"{baseName}.tlk.json");

            var jsonToSave = TLKRawModeCheckBox.IsChecked == true
                ? TLKRawTextBox.Text
                : BuildJsonFromEntries();

            WriteTlkJsonFile(tempJson, jsonToSave);

            var success = RunNwnTlk(
                toolPath,
                $"-i \"{tempJson}\" -l json -o \"{outputPath}\" -k tlk --quiet",
                "save json");

            if (success)
            {
                _currentTlkPath = outputPath;
                FileInfoTextBlock.Text = outputPath;
                SetStatus("TLK saved");
            }
        }

        private void ExportTlkToBinary()
        {
            var toolPath = ResolveNwnTlkPath();
            if (toolPath == null)
            {
                return;
            }

            var saveDialog = new SaveFileDialog
            {
                Filter = "TLK files (*.tlk)|*.tlk|All files (*.*)|*.*",
                Title = "Export TLK Binary",
                FileName = _currentTlkPath != null
                    ? Path.GetFileName(_currentTlkPath)
                    : "dialog.tlk"
            };

            if (saveDialog.ShowDialog() != true)
            {
                return;
            }

            var outputPath = saveDialog.FileName;
            var baseName = Path.GetFileNameWithoutExtension(outputPath);
            var tempJson = Path.Combine(Path.GetTempPath(), $"{baseName}.tlk.json");

            var jsonToSave = TLKRawModeCheckBox.IsChecked == true
                ? TLKRawTextBox.Text
                : BuildJsonFromEntries();

            WriteTlkJsonFile(tempJson, jsonToSave);

            var success = RunNwnTlk(
                toolPath,
                $"-i \"{tempJson}\" -l json -o \"{outputPath}\" -k tlk --quiet",
                "export binary");

            if (success)
            {
                _currentTlkPath = outputPath;
                FileInfoTextBlock.Text = outputPath;
                SetStatus("TLK binary exported");
                LogMessage($"Exported TLK binary: {outputPath}");
            }
        }

        private string? ResolveNwnTlkPath()
        {
            if (!string.IsNullOrWhiteSpace(_nwnTlkPath) && File.Exists(_nwnTlkPath))
            {
                return _nwnTlkPath;
            }

            var candidate = TryResolveToolPath("nwn_tlk.exe");
            if (candidate != null)
            {
                _nwnTlkPath = candidate;
                return _nwnTlkPath;
            }

            var dialog = new OpenFileDialog
            {
                Filter = "nwn_tlk.exe|nwn_tlk.exe|Executable files (*.exe)|*.exe|All files (*.*)|*.*",
                Title = "Locate nwn_tlk.exe"
            };

            if (dialog.ShowDialog() == true)
            {
                if (!dialog.FileName.EndsWith("nwn_tlk.exe", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("Please select nwn_tlk.exe (not a TLK file).", "Aribeth",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return null;
                }

                _nwnTlkPath = dialog.FileName;
                return _nwnTlkPath;
            }

            SetStatus("nwn_tlk.exe not found");
            return null;
        }

        private bool RunNwnTlk(string toolPath, string args, string label)
        {
            LogMessage($"Running: {toolPath} {args}");
            var (exitCode, output) = RunProcess(toolPath, args);
            if (!string.IsNullOrWhiteSpace(output))
            {
                LogMessage(output);
            }

            if (exitCode == 0)
            {
                return true;
            }

            LogMessage($"{label} failed (exit code {exitCode}).");
            SetStatus($"{label} failed. See log for details.");
            return false;
        }

        private void LoadTlkEntriesFromText(string rawText)
        {
            var entries = ParseTlkEntriesFromText(rawText);
            LoadTlkEntries(entries);
        }

        private static List<TlkEntry> ParseTlkEntriesFromText(string rawText)
        {
            var entries = new List<TlkEntry>();
            var lines = rawText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
            var index = TlkBaseIndex;
            TlkEntry? current = null;

            foreach (var line in lines)
            {
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                var trimmed = line.TrimStart();
                if (trimmed.StartsWith("#", StringComparison.Ordinal) ||
                    trimmed.StartsWith("//", StringComparison.Ordinal))
                {
                    continue;
                }

                if (trimmed.StartsWith("|", StringComparison.Ordinal))
                {
                    if (current != null)
                    {
                        var continuation = trimmed.Substring(1).TrimStart();
                        current.Text = string.Concat(current.Text ?? string.Empty,
                            string.IsNullOrEmpty(current.Text) ? string.Empty : "\n",
                            continuation);
                        current.RecalculateLength();
                    }

                    continue;
                }

                var entry = ParseTlkLine(line, index);
                entry.RecalculateLength();
                entries.Add(entry);
                current = entry;
                index = entry.Index + 1;
            }

            return entries;
        }

        private void LoadTlkEntries(List<TlkEntry> entries, string? jsonOverride = null)
        {
            _tlkEntries = entries;
            DataGridTLK.ItemsSource = _tlkEntries;

            _tlkJsonCache = jsonOverride ?? JsonSerializer.Serialize(_tlkEntries, new JsonSerializerOptions
            {
                WriteIndented = true
            });
            SetStatus("TLK converted to JSON and loaded");
        }

        private static TlkEntry ParseTlkLine(string line, int fallbackIndex)
        {
            var pipeIndex = line.IndexOf('|');
            if (pipeIndex > 0 && int.TryParse(line.Substring(0, pipeIndex).Trim(), out var pipeId))
            {
                var text = line.Substring(pipeIndex + 1);
                return NormalizeTlkText(new TlkEntry
                {
                    Index = pipeId,
                    Text = text.TrimStart()
                });
            }

            var tabParts = line.Split('\t');
            if (tabParts.Length >= 4 && int.TryParse(tabParts[0].Trim(), out var index))
            {
                var entry = new TlkEntry
                {
                    Index = index,
                    SoundResRef = tabParts[1].Trim(),
                    Duration = tabParts[2].Trim(),
                    Text = string.Join("\t", tabParts.Skip(3)).Trim()
                };
                return NormalizeTlkText(StripIndexPrefix(entry));
            }

            var pipeParts = line.Split('|').Select(part => part.Trim()).ToArray();
            if (pipeParts.Length >= 4 && int.TryParse(pipeParts[0], out index))
            {
                var entry = new TlkEntry
                {
                    Index = index,
                    SoundResRef = pipeParts[1],
                    Duration = pipeParts[2],
                    Text = string.Join(" | ", pipeParts.Skip(3))
                };
                return NormalizeTlkText(StripIndexPrefix(entry));
            }

            if (pipeParts.Length >= 2 && int.TryParse(pipeParts[0], out index))
            {
                var entry = new TlkEntry
                {
                    Index = index,
                    Text = string.Join(" | ", pipeParts.Skip(1))
                };
                return NormalizeTlkText(StripIndexPrefix(entry));
            }

            var whitespaceMatch = Regex.Match(line, @"^\s*(\d+)\s+(.*)$");
            if (whitespaceMatch.Success && int.TryParse(whitespaceMatch.Groups[1].Value, out index))
            {
                return NormalizeTlkText(new TlkEntry
                {
                    Index = index,
                    Text = whitespaceMatch.Groups[2].Value.Trim()
                });
            }

            return NormalizeTlkText(StripIndexPrefix(new TlkEntry
            {
                Index = fallbackIndex,
                Text = line.Trim()
            }));
        }

        private static TlkEntry StripIndexPrefix(TlkEntry entry)
        {
            if (string.IsNullOrEmpty(entry.Text))
            {
                return entry;
            }

            var indexPrefix = entry.Index.ToString();
            if (entry.Text.StartsWith(indexPrefix + "|", StringComparison.Ordinal))
            {
                entry.Text = entry.Text.Substring(indexPrefix.Length + 1).TrimStart();
            }
            else if (entry.Text.StartsWith(indexPrefix + "\t", StringComparison.Ordinal))
            {
                entry.Text = entry.Text.Substring(indexPrefix.Length + 1).TrimStart();
            }
            else if (entry.Text.StartsWith(indexPrefix + " ", StringComparison.Ordinal))
            {
                entry.Text = entry.Text.Substring(indexPrefix.Length + 1).TrimStart();
            }

            return entry;
        }

        private static TlkEntry NormalizeTlkText(TlkEntry entry)
        {
            if (string.IsNullOrEmpty(entry.Text))
            {
                return entry;
            }

            var text = entry.Text.TrimStart();
            while (text.StartsWith("|", StringComparison.Ordinal))
            {
                text = text.Substring(1).TrimStart();
            }

            entry.Text = text;
            return entry;
        }

        private string SerializeEntriesToText(bool useBaseIndices)
        {
            var builder = new StringBuilder();
            var ordered = _tlkEntries
                .Where(entry => entry != null)
                .GroupBy(entry => entry.Index)
                .Select(group => group.Last())
                .OrderBy(entry => entry.Index)
                .ToList();

            foreach (var entry in ordered)
            {
                entry.RecalculateLength();
                var index = useBaseIndices && entry.Index >= TlkUserIndexBase
                    ? entry.Index - TlkUserIndexBase
                    : entry.Index;
                builder.Append(index);
                builder.Append('|');

                var text = entry.Text ?? string.Empty;
                var normalized = text.Replace("\r\n", "\n").Replace("\r", "\n");
                var lines = normalized.Split('\n');
                builder.AppendLine(lines.Length > 0 ? lines[0] : string.Empty);

                const string continuationPrefix = "        |";
                for (var i = 1; i < lines.Length; i++)
                {
                    builder.AppendLine($"{continuationPrefix}{lines[i]}");
                }
            }

            return builder.ToString();
        }

        private bool RunNwnTlkImportWithFallbacks(string toolPath, string outputPath, string userTextPath, string baseTextPath)
        {
            var attempts = new[]
            {
                ($"-i \"{userTextPath}\" -o \"{outputPath}\" -j text -k tlk -u", "import user"),
                ($"-i \"{baseTextPath}\" -o \"{outputPath}\" -j text -k tlk -b", "import base"),
                ($"-i \"{userTextPath}\" -o \"{outputPath}\" -j text -k tlk", "import user (no flag)"),
                ($"-i \"{baseTextPath}\" -o \"{outputPath}\" -j text -k tlk", "import base (no flag)")
            };

            foreach (var (args, label) in attempts)
            {
                LogMessage($"Running: {toolPath} {args}");
                var (exitCode, output) = RunProcess(toolPath, args);
                if (!string.IsNullOrWhiteSpace(output))
                {
                    LogMessage(output);
                }

                if (exitCode == 0)
                {
                    return true;
                }

                LogMessage($"{label} attempt failed (exit code {exitCode}).");
            }

            SetStatus("TLK import failed. See log for details.");
            return false;
        }

        private static string ConvertUserToBaseIndices(string text)
        {
            var lines = text.Replace("\r\n", "\n").Replace("\r", "\n").Split('\n');
            for (var i = 0; i < lines.Length; i++)
            {
                var line = lines[i];
                if (string.IsNullOrWhiteSpace(line))
                {
                    continue;
                }

                if (line.TrimStart().StartsWith("|", StringComparison.Ordinal))
                {
                    continue;
                }

                var pipeIndex = line.IndexOf('|');
                if (pipeIndex > 0 && int.TryParse(line.Substring(0, pipeIndex).Trim(), out var id))
                {
                    if (id >= TlkUserIndexBase)
                    {
                        lines[i] = (id - TlkUserIndexBase) + line.Substring(pipeIndex);
                    }
                }
            }

            return string.Join("\n", lines);
        }
        private static void WriteTlkJsonFile(string path, string content)
        {
            File.WriteAllText(path, content, new UTF8Encoding(encoderShouldEmitUTF8Identifier: false));
        }

        private List<TlkEntry> ParseTlkEntriesFromJson(string rawJson)
        {
            _tlkJsonRoot = JsonNode.Parse(rawJson) as JsonObject;
            var entriesNode = _tlkJsonRoot?["entries"] as JsonArray;
            var entries = new List<TlkEntry>();

            if (entriesNode == null)
            {
                return entries;
            }

            var entryMap = new Dictionary<int, TlkEntry>();
            var maxId = -1;

            for (var i = 0; i < entriesNode.Count; i++)
            {
                var entryNode = entriesNode[i];
                if (entryNode is not JsonObject entryObj)
                {
                    maxId = Math.Max(maxId, i);
                    continue;
                }

                var id = entryObj["id"]?.GetValue<int>() ?? i;
                var text = entryObj["text"]?.GetValue<string>() ?? string.Empty;
                var sound = entryObj["sound"]?.GetValue<string>() ?? string.Empty;
                var duration = entryObj["soundLength"]?.ToString() ?? string.Empty;

                var entry = new TlkEntry
                {
                    Index = id + TlkUserIndexBase,
                    Text = text,
                    SoundResRef = sound,
                    Duration = duration
                };
                entry.RecalculateLength();
                entryMap[id] = entry;
                maxId = Math.Max(maxId, id);
            }

            for (var id = 0; id <= maxId; id++)
            {
                if (entryMap.TryGetValue(id, out var entry))
                {
                    entries.Add(entry);
                    continue;
                }

                entries.Add(new TlkEntry
                {
                    Index = id + TlkUserIndexBase,
                    Text = string.Empty,
                    SoundResRef = string.Empty,
                    Duration = string.Empty,
                    Length = 0
                });
            }

            return entries;
        }

        private string BuildJsonFromEntries()
        {
            if (_tlkJsonRoot == null)
            {
                _tlkJsonRoot = new JsonObject
                {
                    ["language"] = 0,
                    ["entries"] = new JsonArray()
                };
            }

            var entriesArray = _tlkJsonRoot["entries"] as JsonArray ?? new JsonArray();
            _tlkJsonRoot["entries"] = entriesArray;

            var existingMap = new Dictionary<int, JsonObject>();
            foreach (var node in entriesArray.OfType<JsonObject>())
            {
                try
                {
                    var idNode = node["id"];
                    if (idNode != null)
                    {
                        var id = idNode.GetValue<int>();
                        existingMap[id] = node;
                    }
                }
                catch
                {
                    // Skip malformed entries.
                }
            }

            var ordered = _tlkEntries
                .Where(entry => entry != null)
                .GroupBy(entry => entry.Index)
                .Select(group => group.Last())
                .OrderBy(entry => entry.Index)
                .ToList();

            var newArray = new JsonArray();
            foreach (var entry in ordered)
            {
                var id = entry.Index >= TlkUserIndexBase ? entry.Index - TlkUserIndexBase : entry.Index;
                var obj = new JsonObject();
                if (existingMap.TryGetValue(id, out var existing))
                {
                    foreach (var kvp in existing)
                    {
                        if (kvp.Value == null)
                        {
                            obj[kvp.Key] = null;
                            continue;
                        }

                        var cloned = JsonNode.Parse(kvp.Value.ToJsonString());
                        obj[kvp.Key] = cloned;
                    }
                }

                obj["id"] = id;

                obj["text"] = entry.Text ?? string.Empty;

                if (!string.IsNullOrWhiteSpace(entry.SoundResRef))
                {
                    obj["sound"] = entry.SoundResRef;
                }

                if (double.TryParse(entry.Duration, out var soundLength))
                {
                    obj["soundLength"] = soundLength;
                }

                newArray.Add(obj);
            }

            _tlkJsonRoot["entries"] = newArray;
            return _tlkJsonRoot.ToJsonString(new JsonSerializerOptions
            {
                WriteIndented = true
            });
        }

        private void DataGridTLK_CellEditEnding(object? sender, DataGridCellEditEndingEventArgs e)
        {
            if (e.Row.Item is TlkEntry entry)
            {
                entry.RecalculateLength();
                CaptureTlkSnapshot(clearRedo: true);
            }
        }


        private void HookTwoDaEvents()
        {
            if (_twoDaTable == null)
            {
                return;
            }

            _twoDaTable.ColumnChanged -= TwoDaTable_ColumnChanged;
            _twoDaTable.RowChanged -= TwoDaTable_RowChanged;
            _twoDaTable.RowDeleted -= TwoDaTable_RowDeleted;
            _twoDaTable.TableNewRow -= TwoDaTable_TableNewRow;
            _twoDaTable.Columns.CollectionChanged -= TwoDaTable_ColumnsChanged;

            _twoDaTable.ColumnChanged += TwoDaTable_ColumnChanged;
            _twoDaTable.RowChanged += TwoDaTable_RowChanged;
            _twoDaTable.RowDeleted += TwoDaTable_RowDeleted;
            _twoDaTable.TableNewRow += TwoDaTable_TableNewRow;
            _twoDaTable.Columns.CollectionChanged += TwoDaTable_ColumnsChanged;
        }

        private void TwoDaTable_ColumnChanged(object? sender, DataColumnChangeEventArgs e) => CaptureTwoDaSnapshot(clearRedo: true);
        private void TwoDaTable_RowChanged(object? sender, DataRowChangeEventArgs e) => CaptureTwoDaSnapshot(clearRedo: true);
        private void TwoDaTable_RowDeleted(object? sender, DataRowChangeEventArgs e) => CaptureTwoDaSnapshot(clearRedo: true);
        private void TwoDaTable_TableNewRow(object? sender, DataTableNewRowEventArgs e) => CaptureTwoDaSnapshot(clearRedo: true);
        private void TwoDaTable_ColumnsChanged(object? sender, CollectionChangeEventArgs e) => CaptureTwoDaSnapshot(clearRedo: true);

        private void CaptureTwoDaSnapshot(bool clearRedo)
        {
            if (_suppressUndoCapture || _twoDaTable == null)
            {
                return;
            }

            var data = BuildTwoDaDataFromTable();
            var serialized = TwoDAParser.Serialize(data);
            if (_twoDaUndo.Count == 0 || _twoDaUndo.Peek() != serialized)
            {
                _twoDaUndo.Push(serialized);
                if (clearRedo)
                {
                    _twoDaRedo.Clear();
                }
            }
        }

        private void CaptureTlkSnapshot(bool clearRedo)
        {
            if (_suppressUndoCapture)
            {
                return;
            }

            var serialized = JsonSerializer.Serialize(_tlkEntries);
            if (_tlkUndo.Count == 0 || _tlkUndo.Peek() != serialized)
            {
                _tlkUndo.Push(serialized);
                if (clearRedo)
                {
                    _tlkRedo.Clear();
                }
            }
        }

        private void Undo()
        {
            if (IsTlkTabActive())
            {
                UndoTlk();
            }
            else
            {
                UndoTwoDa();
            }
        }

        private void Redo()
        {
            if (IsTlkTabActive())
            {
                RedoTlk();
            }
            else
            {
                RedoTwoDa();
            }
        }

        private void UndoTwoDa()
        {
            if (_twoDaUndo.Count <= 1 || _twoDaTable == null)
            {
                SetStatus("Nothing to undo");
                return;
            }

            _suppressUndoCapture = true;
            var current = _twoDaUndo.Pop();
            _twoDaRedo.Push(current);
            var previous = _twoDaUndo.Peek();
            var tempPath = Path.GetTempFileName();
            File.WriteAllText(tempPath, previous);
            var data = TwoDAParser.Parse(tempPath, LogMessage, strictHeader: true);
            UpdateTwoDaFormattingState(data);
            _twoDaTable = ToDataTable(data);
            DataGrid2DA.ItemsSource = _twoDaTable.DefaultView;
            HookTwoDaEvents();
            _suppressUndoCapture = false;
            SetStatus("Undo 2DA");
        }

        private void RedoTwoDa()
        {
            if (_twoDaRedo.Count == 0 || _twoDaTable == null)
            {
                SetStatus("Nothing to redo");
                return;
            }

            _suppressUndoCapture = true;
            var next = _twoDaRedo.Pop();
            _twoDaUndo.Push(next);
            var tempPath = Path.GetTempFileName();
            File.WriteAllText(tempPath, next);
            var data = TwoDAParser.Parse(tempPath, LogMessage, strictHeader: true);
            UpdateTwoDaFormattingState(data);
            _twoDaTable = ToDataTable(data);
            DataGrid2DA.ItemsSource = _twoDaTable.DefaultView;
            HookTwoDaEvents();
            _suppressUndoCapture = false;
            SetStatus("Redo 2DA");
        }

        private void UndoTlk()
        {
            if (_tlkUndo.Count <= 1)
            {
                SetStatus("Nothing to undo");
                return;
            }

            _suppressUndoCapture = true;
            var current = _tlkUndo.Pop();
            _tlkRedo.Push(current);
            var previous = _tlkUndo.Peek();
            RestoreTlkFromJson(previous);
            _suppressUndoCapture = false;
            SetStatus("Undo TLK");
        }

        private void RedoTlk()
        {
            if (_tlkRedo.Count == 0)
            {
                SetStatus("Nothing to redo");
                return;
            }

            _suppressUndoCapture = true;
            var next = _tlkRedo.Pop();
            _tlkUndo.Push(next);
            RestoreTlkFromJson(next);
            _suppressUndoCapture = false;
            SetStatus("Redo TLK");
        }

        private void RestoreTlkFromJson(string json)
        {
            var entries = JsonSerializer.Deserialize<List<TlkEntry>>(json) ?? new List<TlkEntry>();
            _tlkEntries = entries;
            DataGridTLK.ItemsSource = _tlkEntries;
            _tlkJsonCache = json;
        }

        private void Check2DAIssues()
        {
            if (string.IsNullOrWhiteSpace(_currentFilePath))
            {
                SetStatus("No 2DA loaded");
                return;
            }

            var toolPath = ResolveNwn2daPath();
            if (toolPath == null)
            {
                return;
            }

            RunNwn2daWithFallbacks(
                toolPath,
                "check",
                new[]
                {
                    $"check \"{_currentFilePath}\"",
                    $"check -i \"{_currentFilePath}\""
                });
        }

        private void Normalize2DA()
        {
            if (string.IsNullOrWhiteSpace(_currentFilePath))
            {
                SetStatus("No 2DA loaded");
                return;
            }

            var toolPath = ResolveNwn2daPath();
            if (toolPath == null)
            {
                return;
            }

            var saveDialog = new SaveFileDialog
            {
                Filter = "2DA files (*.2da)|*.2da|All files (*.*)|*.*",
                Title = "Save Normalized 2DA",
                FileName = Path.GetFileNameWithoutExtension(_currentFilePath) + ".normalized.2da"
            };

            if (saveDialog.ShowDialog() != true)
            {
                return;
            }

            var outputPath = saveDialog.FileName;
            RunNwn2daWithFallbacks(
                toolPath,
                "normalize",
                new[]
                {
                    $"normalize \"{_currentFilePath}\" \"{outputPath}\"",
                    $"normalize -i \"{_currentFilePath}\" -o \"{outputPath}\""
                },
                () =>
                {
                    LogMessage($"Normalized 2DA saved to: {outputPath}");
                    SetStatus("Normalize complete");
                });
        }

        private void Merge2DA()
        {
            if (string.IsNullOrWhiteSpace(_currentFilePath))
            {
                SetStatus("No 2DA loaded");
                return;
            }

            var toolPath = ResolveNwn2daPath();
            if (toolPath == null)
            {
                return;
            }

            var openDialog = new OpenFileDialog
            {
                Filter = "2DA files (*.2da)|*.2da|All files (*.*)|*.*",
                Title = "Select 2DA to Merge"
            };

            if (openDialog.ShowDialog() != true)
            {
                return;
            }

            var mergePath = openDialog.FileName;
            var saveDialog = new SaveFileDialog
            {
                Filter = "2DA files (*.2da)|*.2da|All files (*.*)|*.*",
                Title = "Save Merged 2DA",
                FileName = Path.GetFileNameWithoutExtension(_currentFilePath) + ".merged.2da"
            };

            if (saveDialog.ShowDialog() != true)
            {
                return;
            }

            var outputPath = saveDialog.FileName;
            RunNwn2daWithFallbacks(
                toolPath,
                "merge",
                new[]
                {
                    $"merge \"{_currentFilePath}\" \"{mergePath}\" \"{outputPath}\"",
                    $"merge -i \"{_currentFilePath}\" -j \"{mergePath}\" -o \"{outputPath}\""
                },
                () =>
                {
                    LogMessage($"Merged 2DA saved to: {outputPath}");
                    SetStatus("Merge complete");
                });
        }

        private string? ResolveNwn2daPath()
        {
            if (!string.IsNullOrWhiteSpace(_nwn2daPath) && File.Exists(_nwn2daPath))
            {
                return _nwn2daPath;
            }

            var candidate = TryResolveToolPath("nwn-2da.exe");
            if (candidate != null)
            {
                _nwn2daPath = candidate;
                return _nwn2daPath;
            }

            var dialog = new OpenFileDialog
            {
                Filter = "nwn-2da.exe|nwn-2da.exe|Executable files (*.exe)|*.exe|All files (*.*)|*.*",
                Title = "Locate nwn-2da.exe"
            };

            if (dialog.ShowDialog() == true)
            {
                if (!dialog.FileName.EndsWith("nwn-2da.exe", StringComparison.OrdinalIgnoreCase))
                {
                    MessageBox.Show("Please select nwn-2da.exe.", "Aribeth",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                    return null;
                }

                _nwn2daPath = dialog.FileName;
                return _nwn2daPath;
            }

            SetStatus("nwn-2da.exe not found");
            return null;
        }

        private static string? TryResolveToolPath(string exeName)
        {
            var baseDir = AppDomain.CurrentDomain.BaseDirectory;
            var candidates = new[]
            {
                Path.Combine(baseDir, exeName),
                Path.Combine(baseDir, "plugins", exeName),
                Path.GetFullPath(Path.Combine(baseDir, "..", "..", "..", "plugins", exeName))
            };

            foreach (var candidate in candidates)
            {
                if (File.Exists(candidate))
                {
                    return candidate;
                }
            }

            return null;
        }

        private void RunNwn2daWithFallbacks(string toolPath, string verb, string[] attempts, Action? onSuccess = null)
        {
            foreach (var args in attempts)
            {
                LogMessage($"Running: {toolPath} {args}");
                var (exitCode, output) = RunProcess(toolPath, args);
                LogMessage(output);
                if (exitCode == 0)
                {
                    onSuccess?.Invoke();
                    return;
                }

                LogMessage($"{verb} attempt failed (exit code {exitCode}).");
            }

            SetStatus($"{verb} failed. See log for details.");
        }

        private static (int ExitCode, string Output) RunProcess(string fileName, string args)
        {
            var process = new Process
            {
                StartInfo = new ProcessStartInfo
                {
                    FileName = fileName,
                    Arguments = args,
                    UseShellExecute = false,
                    CreateNoWindow = true,
                    RedirectStandardOutput = true,
                    RedirectStandardError = true
                }
            };

            process.Start();
            var output = process.StandardOutput.ReadToEnd();
            var error = process.StandardError.ReadToEnd();
            process.WaitForExit();

            var combined = new StringBuilder();
            if (!string.IsNullOrWhiteSpace(output))
            {
                combined.AppendLine(output.TrimEnd());
            }

            if (!string.IsNullOrWhiteSpace(error))
            {
                combined.AppendLine(error.TrimEnd());
            }

            return (process.ExitCode, combined.ToString().Trim());
        }

        private void ApplyColumnFilter()
        {
            if (_twoDaTable == null)
            {
                return;
            }

            var filterText = ColumnFilterTextBox.Text.Trim();
            Regex? regex = null;
            if (!string.IsNullOrEmpty(filterText))
            {
                try
                {
                    regex = new Regex(filterText, RegexOptions.IgnoreCase);
                }
                catch (ArgumentException)
                {
                    SetStatus("Invalid column filter regex");
                    return;
                }
            }

            foreach (var column in DataGrid2DA.Columns)
            {
                if (column.Header == null)
                {
                    continue;
                }

                var header = column.Header.ToString() ?? string.Empty;
                if (header == "Row")
                {
                    column.Visibility = Visibility.Visible;
                    continue;
                }

                if (regex == null)
                {
                    column.Visibility = Visibility.Visible;
                }
                else
                {
                    column.Visibility = regex.IsMatch(header) ? Visibility.Visible : Visibility.Collapsed;
                }
            }

            SetStatus("Column filter updated");
        }

        private void ApplyRowFilter()
        {
            if (_twoDaTable == null)
            {
                return;
            }

            var searchText = SearchTextBox.Text;
            if (string.IsNullOrWhiteSpace(searchText))
            {
                _twoDaTable.DefaultView.RowFilter = string.Empty;
                SetStatus("Search cleared");
                return;
            }

            if (RegexCheckBox.IsChecked == true)
            {
                try
                {
                    _ = new Regex(searchText, RegexOptions.IgnoreCase);
                }
                catch (ArgumentException)
                {
                    SetStatus("Invalid search regex");
                    return;
                }

                // DataView.RowFilter does not support regex; keep all rows and rely on selection.
                _twoDaTable.DefaultView.RowFilter = string.Empty;
            }
            else if (WildcardCheckBox.IsChecked == true)
            {
                var likePattern = BuildLikePattern(searchText).Replace("'", "''");
                var conditions = _twoDaTable.Columns
                    .Cast<DataColumn>()
                    .Select(column => $"CONVERT([{column.ColumnName}], 'System.String') LIKE '{likePattern}'");
                _twoDaTable.DefaultView.RowFilter = string.Join(" OR ", conditions);
            }
            else
            {
                var escaped = searchText.Replace("'", "''");
                var conditions = _twoDaTable.Columns
                    .Cast<DataColumn>()
                    .Select(column => $"CONVERT([{column.ColumnName}], 'System.String') LIKE '%{escaped}%'");
                _twoDaTable.DefaultView.RowFilter = string.Join(" OR ", conditions);
            }

            SetStatus("Search updated");
        }

        private void SelectFirstMatchIn2da()
        {
            if (_twoDaTable == null)
            {
                return;
            }

            var query = SearchTextBox.Text;
            if (string.IsNullOrWhiteSpace(query))
            {
                return;
            }

            Regex? regex = null;
            Regex? wildcardRegex = null;
            if (RegexCheckBox.IsChecked == true)
            {
                try
                {
                    regex = new Regex(query, RegexOptions.IgnoreCase);
                }
                catch
                {
                    return;
                }
            }
            else if (WildcardCheckBox.IsChecked == true)
            {
                wildcardRegex = BuildWildcardRegex(query, ignoreCase: true);
            }

            foreach (DataRowView rowView in _twoDaTable.DefaultView)
            {
                foreach (DataColumn column in _twoDaTable.Columns)
                {
                    var cell = rowView.Row[column]?.ToString() ?? string.Empty;
                    var match = regex != null
                        ? regex.IsMatch(cell)
                        : wildcardRegex != null
                            ? wildcardRegex.IsMatch(cell)
                            : cell.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0;

                    if (match)
                    {
                        DataGrid2DA.SelectedCells.Clear();
                        DataGrid2DA.SelectedItem = rowView;
                        var targetColumn = DataGrid2DA.Columns.FirstOrDefault(c => c.Header?.ToString() == column.ColumnName);
                        DataGrid2DA.ScrollIntoView(rowView, targetColumn);
                        DataGrid2DA.CurrentCell = new DataGridCellInfo(rowView, targetColumn);
                        return;
                    }
                }
            }
        }

        private void FindNextMatchIn2da(bool forward)
        {
            if (_twoDaTable == null)
            {
                return;
            }

            var query = SearchTextBox.Text;
            if (string.IsNullOrWhiteSpace(query))
            {
                return;
            }

            Regex? regex = null;
            Regex? wildcardRegex = null;
            if (RegexCheckBox.IsChecked == true)
            {
                try
                {
                    regex = new Regex(query, RegexOptions.IgnoreCase);
                }
                catch
                {
                    return;
                }
            }
            else if (WildcardCheckBox.IsChecked == true)
            {
                wildcardRegex = BuildWildcardRegex(query, ignoreCase: true);
            }

            var view = _twoDaTable.DefaultView;
            var rows = view.Cast<DataRowView>().ToList();
            if (rows.Count == 0)
            {
                return;
            }

            var startRowIndex = 0;
            var startColIndex = 0;
            if (DataGrid2DA.CurrentItem is DataRowView currentRowView)
            {
                startRowIndex = rows.IndexOf(currentRowView);
                startColIndex = Math.Max(0, DataGrid2DA.Columns.IndexOf(DataGrid2DA.CurrentCell.Column));
            }

            var columns = _twoDaTable.Columns.Cast<DataColumn>().ToList();
            var total = rows.Count * columns.Count;
            for (var step = 1; step <= total; step++)
            {
                var offset = forward ? step : -step;
                var flatIndex = (startRowIndex * columns.Count + startColIndex + offset + total) % total;
                var rowIndex = flatIndex / columns.Count;
                var colIndex = flatIndex % columns.Count;

                var rowView = rows[rowIndex];
                var column = columns[colIndex];
                var cell = rowView.Row[column]?.ToString() ?? string.Empty;
                var match = regex != null
                    ? regex.IsMatch(cell)
                    : wildcardRegex != null
                        ? wildcardRegex.IsMatch(cell)
                        : cell.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0;

                if (match)
                {
                    DataGrid2DA.SelectedCells.Clear();
                    DataGrid2DA.SelectedItem = rowView;
                    var targetColumn = DataGrid2DA.Columns.FirstOrDefault(c => c.Header?.ToString() == column.ColumnName);
                    DataGrid2DA.ScrollIntoView(rowView, targetColumn);
                    DataGrid2DA.CurrentCell = new DataGridCellInfo(rowView, targetColumn);
                    return;
                }
            }
        }

        private void SelectFirstMatchInTlk()
        {
            var query = TLKSearchTextBox.Text;
            if (string.IsNullOrWhiteSpace(query))
            {
                return;
            }

            Regex? regex = null;
            if (TLKRegexCheckBox.IsChecked == true)
            {
                try
                {
                    regex = new Regex(query, TLKCaseSensitiveCheckBox.IsChecked == true ? RegexOptions.None : RegexOptions.IgnoreCase);
                }
                catch
                {
                    return;
                }
            }

            var comparison = TLKCaseSensitiveCheckBox.IsChecked == true
                ? StringComparison.Ordinal
                : StringComparison.OrdinalIgnoreCase;

            foreach (var entry in _tlkEntries)
            {
                var text = entry.Text ?? string.Empty;
                var match = regex != null
                    ? regex.IsMatch(text)
                    : text.IndexOf(query, comparison) >= 0;

                if (match)
                {
                    DataGridTLK.SelectedItem = entry;
                    DataGridTLK.ScrollIntoView(entry);
                    return;
                }
            }
        }

        private void FindNextMatchInTlk(bool forward)
        {
            var query = TLKSearchTextBox.Text;
            if (string.IsNullOrWhiteSpace(query))
            {
                return;
            }

            Regex? regex = null;
            if (TLKRegexCheckBox.IsChecked == true)
            {
                try
                {
                    regex = new Regex(query, TLKCaseSensitiveCheckBox.IsChecked == true ? RegexOptions.None : RegexOptions.IgnoreCase);
                }
                catch
                {
                    return;
                }
            }

            var comparison = TLKCaseSensitiveCheckBox.IsChecked == true
                ? StringComparison.Ordinal
                : StringComparison.OrdinalIgnoreCase;

            if (_tlkEntries.Count == 0)
            {
                return;
            }

            var startIndex = 0;
            if (DataGridTLK.SelectedItem is TlkEntry selected)
            {
                startIndex = _tlkEntries.IndexOf(selected);
                if (startIndex < 0)
                {
                    startIndex = 0;
                }
            }

            var total = _tlkEntries.Count;
            for (var step = 1; step <= total; step++)
            {
                var offset = forward ? step : -step;
                var index = (startIndex + offset + total) % total;
                var entry = _tlkEntries[index];
                var text = entry.Text ?? string.Empty;
                var match = regex != null
                    ? regex.IsMatch(text)
                    : text.IndexOf(query, comparison) >= 0;

                if (match)
                {
                    DataGridTLK.SelectedItem = entry;
                    DataGridTLK.ScrollIntoView(entry);
                    return;
                }
            }
        }

        private void GoToTlkIndex()
        {
            if (_tlkEntries.Count == 0)
            {
                SetStatus("No TLK loaded");
                return;
            }

            var raw = TLKIndexTextBox.Text?.Trim();
            if (string.IsNullOrWhiteSpace(raw))
            {
                SetStatus("Enter a TLK index");
                return;
            }

            if (!int.TryParse(raw, out var index))
            {
                SetStatus("Invalid TLK index");
                return;
            }

            var entry = _tlkEntries.FirstOrDefault(e => e.Index == index);
            if (entry == null && index < TlkUserIndexBase)
            {
                entry = _tlkEntries.FirstOrDefault(e => e.Index == index + TlkUserIndexBase);
                if (entry != null)
                {
                    index = entry.Index;
                }
            }

            if (entry == null)
            {
                SetStatus($"Index {index} not found");
                return;
            }

            if (TLKRawModeCheckBox.IsChecked == true)
            {
                TLKRawModeCheckBox.IsChecked = false;
            }

            DataGridTLK.SelectedItem = entry;
            var targetColumn = DataGridTLK.Columns.FirstOrDefault(c => string.Equals(c.Header?.ToString(), "Text", StringComparison.OrdinalIgnoreCase));
            DataGridTLK.ScrollIntoView(entry, targetColumn);
            DataGridTLK.CurrentCell = new DataGridCellInfo(entry, targetColumn);
            SetStatus($"Navigated to index {index}");
        }

        private void ClearFilters()
        {
            ColumnFilterTextBox.Text = string.Empty;
            SearchTextBox.Text = string.Empty;
            RegexCheckBox.IsChecked = false;
            ApplyColumnFilter();
            ApplyRowFilter();
            SetStatus("Filters cleared");
        }

        private void CopyRows(bool cut)
        {
            if (_twoDaTable == null)
            {
                return;
            }

            var selected = GetSelectedRowViews();
            if (selected.Count == 0)
            {
                SetStatus("No rows selected");
                return;
            }

            _rowClipboard.Clear();
            foreach (var rowView in selected)
            {
                var values = rowView.Row.ItemArray.Select(value => value?.ToString() ?? string.Empty).ToList();
                _rowClipboard.Add(values);
            }

            var text = new StringBuilder();
            foreach (var row in _rowClipboard)
            {
                text.AppendLine(string.Join("\t", row));
            }

            Clipboard.SetText(text.ToString());
            SetStatus(cut ? "Cut rows" : "Copied rows");

            if (cut)
            {
                foreach (var rowView in selected)
                {
                    rowView.Row.Delete();
                }

                _twoDaTable.AcceptChanges();
            }
        }

        private void PasteRows()
        {
            if (_twoDaTable == null)
            {
                return;
            }

            var insertionIndex = GetInsertionIndex();
            if (_rowClipboard.Count == 0 && Clipboard.ContainsText())
            {
                _rowClipboard.AddRange(ParseClipboardRows(Clipboard.GetText(), _twoDaTable.Columns.Count));
            }

            if (_rowClipboard.Count == 0)
            {
                SetStatus("Clipboard is empty");
                return;
            }

            foreach (var rowValues in _rowClipboard)
            {
                var row = _twoDaTable.NewRow();
                for (var i = 0; i < _twoDaTable.Columns.Count; i++)
                {
                    row[i] = i < rowValues.Count ? rowValues[i] : "****";
                }

                if (insertionIndex >= 0 && insertionIndex < _twoDaTable.Rows.Count)
                {
                    _twoDaTable.Rows.InsertAt(row, insertionIndex++);
                }
                else
                {
                    _twoDaTable.Rows.Add(row);
                }
            }

            _twoDaTable.AcceptChanges();
            SetStatus("Pasted rows");
        }

        private static List<List<string>> ParseClipboardRows(string clipboardText, int columnCount)
        {
            var rows = new List<List<string>>();
            var lines = clipboardText.Split(new[] { "\r\n", "\n" }, StringSplitOptions.RemoveEmptyEntries);
            foreach (var line in lines)
            {
                var tokens = line.Split('\t').Select(token => token.Trim()).ToList();
                while (tokens.Count < columnCount)
                {
                    tokens.Add("****");
                }

                rows.Add(tokens);
            }

            return rows;
        }

        private void InsertRow(bool above)
        {
            if (_twoDaTable == null)
            {
                return;
            }

            var insertionIndex = GetInsertionIndex();
            if (!above)
            {
                insertionIndex++;
            }

            var newRow = _twoDaTable.NewRow();
            for (var i = 0; i < _twoDaTable.Columns.Count; i++)
            {
                newRow[i] = i == 0 ? string.Empty : "****";
            }

            if (insertionIndex < 0 || insertionIndex > _twoDaTable.Rows.Count)
            {
                _twoDaTable.Rows.Add(newRow);
            }
            else
            {
                _twoDaTable.Rows.InsertAt(newRow, insertionIndex);
            }

            _twoDaTable.AcceptChanges();
            SetStatus(above ? "Inserted row above" : "Inserted row below");
        }

        private void DeleteSelectedRows()
        {
            if (_twoDaTable == null)
            {
                return;
            }

            var selected = GetSelectedRowViews();
            if (selected.Count == 0)
            {
                SetStatus("No rows selected");
                return;
            }

            foreach (var rowView in selected)
            {
                rowView.Row.Delete();
            }

            _twoDaTable.AcceptChanges();
            SetStatus("Deleted rows");
        }

        private List<DataRowView> GetSelectedRowViews()
        {
            return DataGrid2DA.SelectedItems
                .OfType<DataRowView>()
                .ToList();
        }

        private int GetInsertionIndex()
        {
            if (DataGrid2DA.CurrentItem is DataRowView rowView)
            {
                return rowView.Row.Table.Rows.IndexOf(rowView.Row);
            }

            return _twoDaTable?.Rows.Count ?? -1;
        }

        private void AddColumn()
        {
            if (_twoDaTable == null)
            {
                SetStatus("No 2DA loaded");
                return;
            }

            var name = PromptForText("Add Column", "Column name:");
            if (string.IsNullOrWhiteSpace(name))
            {
                return;
            }

            if (_twoDaTable.Columns.Contains(name))
            {
                SetStatus("Column already exists");
                return;
            }

            _twoDaTable.Columns.Add(name, typeof(string));
            foreach (DataRow row in _twoDaTable.Rows)
            {
                row[name] = "****";
            }

            _twoDaColumnWidths.Add(Math.Max(1, name.Length));
            _twoDaTable.AcceptChanges();
            SetStatus($"Added column: {name}");
        }

        private void RemoveColumn()
        {
            if (_twoDaTable == null)
            {
                SetStatus("No 2DA loaded");
                return;
            }

            var columns = _twoDaTable.Columns
                .Cast<DataColumn>()
                .Select(column => column.ColumnName)
                .Where(name => name != "Row")
                .ToList();

            if (columns.Count == 0)
            {
                SetStatus("No columns to remove");
                return;
            }

            var selected = PromptForSelection("Remove Column", "Select a column to remove:", columns);
            if (string.IsNullOrWhiteSpace(selected))
            {
                return;
            }

            var removeIndex = columns.IndexOf(selected);
            _twoDaTable.Columns.Remove(selected);
            if (removeIndex >= 0 && removeIndex < _twoDaColumnWidths.Count)
            {
                _twoDaColumnWidths.RemoveAt(removeIndex);
            }
            _twoDaTable.AcceptChanges();
            SetStatus($"Removed column: {selected}");
        }

        private void ReplaceValues()
        {
            if (_twoDaTable == null)
            {
                SetStatus("No 2DA loaded");
                return;
            }

            var findText = SearchTextBox.Text ?? string.Empty;
            var replaceText = ReplaceTextBox.Text ?? string.Empty;
            if (string.IsNullOrWhiteSpace(findText))
            {
                SetStatus("Find text is empty");
                return;
            }

            var targets = _twoDaTable.DefaultView.Cast<DataRowView>().ToList();

            if (targets.Count == 0)
            {
                SetStatus("No rows to replace");
                return;
            }

            var replacements = 0;
            if (RegexCheckBox.IsChecked == true)
            {
                Regex regex;
                try
                {
                    regex = new Regex(findText, RegexOptions.IgnoreCase);
                }
                catch (ArgumentException)
                {
                    SetStatus("Invalid regex");
                    return;
                }

                foreach (var rowView in targets)
                {
                    foreach (DataColumn column in _twoDaTable.Columns)
                    {
                        if (column.ColumnName == "Row")
                        {
                            continue;
                        }

                        var original = rowView.Row[column]?.ToString() ?? string.Empty;
                        var updated = regex.Replace(original, replaceText);
                        if (!string.Equals(original, updated, StringComparison.Ordinal))
                        {
                            rowView.Row[column] = updated;
                            replacements++;
                        }
                    }
                }
            }
            else if (WildcardCheckBox.IsChecked == true)
            {
                var wildcardRegex = BuildWildcardRegex(findText, ignoreCase: true);
                foreach (var rowView in targets)
                {
                    foreach (DataColumn column in _twoDaTable.Columns)
                    {
                        if (column.ColumnName == "Row")
                        {
                            continue;
                        }

                        var original = rowView.Row[column]?.ToString() ?? string.Empty;
                        var updated = wildcardRegex.Replace(original, replaceText);
                        if (!string.Equals(original, updated, StringComparison.Ordinal))
                        {
                            rowView.Row[column] = updated;
                            replacements++;
                        }
                    }
                }
            }
            else
            {
                var comparison = StringComparison.OrdinalIgnoreCase;
                foreach (var rowView in targets)
                {
                    foreach (DataColumn column in _twoDaTable.Columns)
                    {
                        if (column.ColumnName == "Row")
                        {
                            continue;
                        }

                        var original = rowView.Row[column]?.ToString() ?? string.Empty;
                        var updated = ReplaceText(original, findText, replaceText, comparison);
                        if (!string.Equals(original, updated, StringComparison.Ordinal))
                        {
                            rowView.Row[column] = updated;
                            replacements++;
                        }
                    }
                }
            }

            _twoDaTable.AcceptChanges();
            SetStatus($"Replaced {replacements} value(s)");
        }

        private static string ReplaceText(string input, string findText, string replaceText, StringComparison comparison)
        {
            if (string.IsNullOrEmpty(input))
            {
                return input;
            }

            var builder = new StringBuilder();
            var index = 0;
            while (true)
            {
                var matchIndex = input.IndexOf(findText, index, comparison);
                if (matchIndex < 0)
                {
                    builder.Append(input.AsSpan(index));
                    break;
                }

                builder.Append(input.AsSpan(index, matchIndex - index));
                builder.Append(replaceText);
                index = matchIndex + findText.Length;
            }

            return builder.ToString();
        }

        private static bool ContainsWildcard(string text) => text.Contains('*') || text.Contains('?');

        private static string BuildLikePattern(string pattern)
        {
            var hasWildcard = ContainsWildcard(pattern);
            var builder = new StringBuilder();
            foreach (var c in pattern)
            {
                switch (c)
                {
                    case '*':
                        builder.Append('%');
                        break;
                    case '?':
                        builder.Append('_');
                        break;
                    case '%':
                    case '_':
                    case '[':
                    case ']':
                        builder.Append('[').Append(c).Append(']');
                        break;
                    default:
                        builder.Append(c);
                        break;
                }
            }

            var like = builder.ToString();
            return hasWildcard ? like : $"%{like}%";
        }

        private static Regex BuildWildcardRegex(string pattern, bool ignoreCase)
        {
            var hasWildcard = ContainsWildcard(pattern);
            var escaped = Regex.Escape(pattern)
                .Replace("\\*", ".*")
                .Replace("\\?", ".");
            var regexPattern = hasWildcard ? escaped : $".*{escaped}.*";
            return new Regex(regexPattern, ignoreCase ? RegexOptions.IgnoreCase : RegexOptions.None);
        }

        private string? PromptForText(string title, string label)
        {
            var window = new Window
            {
                Title = title,
                Width = 360,
                Height = 160,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                Owner = this
            };

            var panel = new StackPanel { Margin = new Thickness(12) };
            panel.Children.Add(new TextBlock { Text = label, Margin = new Thickness(0, 0, 0, 8) });
            var textBox = new TextBox { Margin = new Thickness(0, 0, 0, 12) };
            panel.Children.Add(textBox);

            var buttons = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var okButton = new Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            var cancelButton = new Button { Content = "Cancel", Width = 80, IsCancel = true };

            okButton.Click += (_, _) => window.DialogResult = true;
            buttons.Children.Add(okButton);
            buttons.Children.Add(cancelButton);
            panel.Children.Add(buttons);

            window.Content = panel;
            window.ShowInTaskbar = false;

            var result = window.ShowDialog();
            return result == true ? textBox.Text.Trim() : null;
        }

        private string? PromptForSelection(string title, string label, List<string> options)
        {
            var window = new Window
            {
                Title = title,
                Width = 360,
                Height = 260,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                Owner = this
            };

            var panel = new DockPanel { Margin = new Thickness(12) };
            var textBlock = new TextBlock { Text = label, Margin = new Thickness(0, 0, 0, 8) };
            DockPanel.SetDock(textBlock, Dock.Top);
            panel.Children.Add(textBlock);

            var listBox = new ListBox { ItemsSource = options, MinHeight = 120 };
            panel.Children.Add(listBox);

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 8, 0, 0)
            };
            DockPanel.SetDock(buttons, Dock.Bottom);
            var okButton = new Button { Content = "OK", Width = 80, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            var cancelButton = new Button { Content = "Cancel", Width = 80, IsCancel = true };
            okButton.Click += (_, _) => window.DialogResult = true;
            buttons.Children.Add(okButton);
            buttons.Children.Add(cancelButton);
            panel.Children.Add(buttons);

            window.Content = panel;
            window.ShowInTaskbar = false;

            var result = window.ShowDialog();
            return result == true ? listBox.SelectedItem?.ToString() : null;
        }

        private Window CreateReplaceDialog()
        {
            var window = new Window
            {
                Title = "Replace",
                Width = 400,
                Height = 260,
                WindowStartupLocation = WindowStartupLocation.CenterOwner,
                ResizeMode = ResizeMode.NoResize,
                Owner = this
            };

            var panel = new Grid { Margin = new Thickness(12) };
            panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            panel.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });

            var findLabel = new TextBlock { Text = "Find:" };
            panel.Children.Add(findLabel);
            Grid.SetRow(findLabel, 0);

            var findTextBox = new TextBox { Name = "FindTextBox", Margin = new Thickness(0, 4, 0, 8) };
            panel.Children.Add(findTextBox);
            Grid.SetRow(findTextBox, 1);

            var replaceLabel = new TextBlock { Text = "Replace with:" };
            panel.Children.Add(replaceLabel);
            Grid.SetRow(replaceLabel, 2);

            var replaceTextBox = new TextBox { Name = "ReplaceTextBox", Margin = new Thickness(0, 4, 0, 8) };
            panel.Children.Add(replaceTextBox);
            Grid.SetRow(replaceTextBox, 3);

            var optionsPanel = new StackPanel { Orientation = Orientation.Horizontal, Margin = new Thickness(0, 4, 0, 8) };
            var regexCheck = new CheckBox { Name = "RegexCheckBox", Content = "Regex", Margin = new Thickness(0, 0, 8, 0) };
            var caseCheck = new CheckBox { Name = "CaseCheckBox", Content = "Case Sensitive", Margin = new Thickness(0, 0, 8, 0) };
            var selectedOnlyCheck = new CheckBox { Name = "SelectedOnlyCheckBox", Content = "Selected rows only" };
            optionsPanel.Children.Add(regexCheck);
            optionsPanel.Children.Add(caseCheck);
            optionsPanel.Children.Add(selectedOnlyCheck);
            panel.Children.Add(optionsPanel);
            Grid.SetRow(optionsPanel, 4);

            var buttons = new StackPanel
            {
                Orientation = Orientation.Horizontal,
                HorizontalAlignment = HorizontalAlignment.Right,
                Margin = new Thickness(0, 8, 0, 0)
            };
            var okButton = new Button { Content = "Replace", Width = 90, Margin = new Thickness(0, 0, 8, 0), IsDefault = true };
            var cancelButton = new Button { Content = "Cancel", Width = 90, IsCancel = true };
            okButton.Click += (_, _) => window.DialogResult = true;
            buttons.Children.Add(okButton);
            buttons.Children.Add(cancelButton);

            var wrapper = new DockPanel();
            DockPanel.SetDock(buttons, Dock.Bottom);
            wrapper.Children.Add(buttons);
            wrapper.Children.Add(panel);

            window.Content = wrapper;
            window.RegisterName(findTextBox.Name, findTextBox);
            window.RegisterName(replaceTextBox.Name, replaceTextBox);
            window.RegisterName(regexCheck.Name, regexCheck);
            window.RegisterName(caseCheck.Name, caseCheck);
            window.RegisterName(selectedOnlyCheck.Name, selectedOnlyCheck);

            window.Tag = findTextBox;
            return window;
        }
    }
}