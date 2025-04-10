using Microsoft.Win32;
using Newtonsoft.Json;
using OfficeOpenXml;
using System.IO;
using System.Windows;

namespace ExcelToJson
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private string _filePath = string.Empty;

        public MainWindow()
        {
            InitializeComponent();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
        }

        private void Grid_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string filePath = files[0];
                    if (Path.GetExtension(filePath).Equals(".xlsx", StringComparison.OrdinalIgnoreCase) ||
                        Path.GetExtension(filePath).Equals(".xls", StringComparison.OrdinalIgnoreCase))
                    {
                        var headers = GetExcelHeaders(filePath);
                        var fieldMappings = GenerateDefaultMappings(headers);

                        TextBox_JsonSchema.Text = JsonConvert.SerializeObject(fieldMappings, Formatting.Indented);
                        Button_Save.IsEnabled = true;
                        _filePath = filePath;
                    }
                    else
                    {
                        MessageBox.Show("Только файлы Excel");
                    }
                }
            }
        }

        private List<string> GetExcelHeaders(string filePath)
        {
            var headers = new List<string>();
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var colCount = worksheet.Dimension.Columns;

                for (int col = 1; col <= colCount; col++)
                {
                    headers.Add(worksheet.Cells[1, col].Text);
                }
            }
            return headers;
        }

        private Dictionary<string, string> GenerateDefaultMappings(List<string> headers)
        {
            var mappings = new Dictionary<string, string>();
            foreach (var header in headers)
            {
                mappings[header] = header;
            }
            return mappings;
        }

        private List<Dictionary<string, object?>> ConvertExcelToJson(string filePath, Dictionary<string, string> mappings)
        {
            var result = new List<Dictionary<string, object?>>();

            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;
                var colCount = worksheet.Dimension.Columns;

                for (int row = 2; row <= rowCount; row++)
                {
                    var rowDict = new Dictionary<string, object?>();

                    for (int col = 1; col <= colCount; col++)
                    {
                        var columnName = worksheet.Cells[1, col].Text;
                        var cellValue = worksheet.Cells[row, col].Text;

                        if (mappings.ContainsKey(columnName))
                        {
                            AddNestedValue(rowDict, mappings[columnName], cellValue);
                        }
                    }

                    result.Add(rowDict);
                }
            }

            return result;
        }

        private void AddNestedValue(Dictionary<string, object?> dict, string path, object? value)
        {
            var parts = path.Split('.');
            var current = dict;

            for (int i = 0; i < parts.Length - 1; i++)
            {
                if (!current!.ContainsKey(parts[i]))
                {
                    current[parts[i]] = new Dictionary<string, object?>();
                }
                current = (Dictionary<string, object?>)current[parts[i]]!;
            }

            if (value is string && (string)value == string.Empty) value = null;

            current[parts.Last()] = value;
        }

        private void Button_Save_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                var fieldMappings = JsonConvert.DeserializeObject<Dictionary<string, string>>(TextBox_JsonSchema.Text)!;
                var jsonData = ConvertExcelToJson(_filePath, fieldMappings);

                SaveJsonToFile(JsonConvert.SerializeObject(jsonData, Formatting.Indented));
            }
            catch
            {
                MessageBox.Show("Schema error");
            }
        }

        private void SaveJsonToFile(string json)
        {
            var saveFileDialog = new SaveFileDialog
            {
                Filter = "JSON files (*.json)|*.json",
                DefaultExt = "json",
                AddExtension = true
            };

            if (saveFileDialog.ShowDialog() == true)
            {
                File.WriteAllText(saveFileDialog.FileName, json);
            }
        }
    }
}