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
                        ProcessExcelFile(filePath);
                    }
                    else
                    {
                        MessageBox.Show("Только файлы Excel");
                    }
                }
            }
        }

        private void ProcessExcelFile(string filePath)
        {
            var file = new FileInfo(filePath);
            using (var package = new ExcelPackage(file))
            {
                var worksheet = package.Workbook.Worksheets[0];
                var rowCount = worksheet.Dimension.Rows;
                var colCount = worksheet.Dimension.Columns;
                var data = new List<Dictionary<string, object>>();

                for (int row = 2; row <= rowCount; row++)
                {
                    var rowDict = new Dictionary<string, object>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        var columnName = worksheet.Cells[1, col].Text;
                        var cellValue = worksheet.Cells[row, col].Text;
                        rowDict[columnName] = cellValue;
                    }
                    data.Add(rowDict);
                }
                var json = JsonConvert.SerializeObject(data, Newtonsoft.Json.Formatting.Indented);
                SaveJsonToFile(json);
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