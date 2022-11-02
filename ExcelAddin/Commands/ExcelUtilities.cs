using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Excel;
using System.Windows;
using System.IO;
using System.Windows.Forms;
using Microsoft.VisualBasic;
using ExcelAddin.Services;

namespace ExcelAddin.Commands
{
    internal static class ExcelUtilities
    {
        private static string _delimiter = "~";

        private static string _startCsvFilePath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        public static void ChangeDelimiter()
        {
            var delimiterFromUser = Interaction.InputBox(
                "Введите разделитель, который будет использоваться при импорте CSV",
                "Разделитель CSV",
                _delimiter);

            if (!String.IsNullOrEmpty(delimiterFromUser)
                && !_delimiter.Equals(delimiterFromUser))
            {
                _delimiter = delimiterFromUser;
            }
        }

        private static string GetCsvPath()
        {
            string csvPath = String.Empty;
            OpenFileDialog openFileDialog = new OpenFileDialog()
            {
                InitialDirectory = _startCsvFilePath,
                Filter = "CSV Файлы (*.csv)|*.csv",
                Multiselect = false,
                RestoreDirectory = true,
                Title = "Выберите csv файл"
            };
            if (openFileDialog.ShowDialog() == DialogResult.OK)
            {
                csvPath = openFileDialog.FileName;
                FileInfo file = new FileInfo(csvPath);
                _startCsvFilePath = file.DirectoryName;
            }
            else
            {
                return String.Empty;
            }
            if (!csvPath.EndsWith(".csv"))
            {
                MessageBox.Show("Неверный формат файла!", "Ошибка");
                return String.Empty;
            }
            return csvPath;
        }

        private static IEnumerable<string[]> GetData(string path)
        {
            var lines = DataService.GetLines(path);
            string[] delimiter = new string[1] { _delimiter };
            return lines.Select(line => line.Split(delimiter, StringSplitOptions.None));
        }


        public static void ImportMultipleCsv()
        {
            Excel.Worksheet activeWorksheet = Globals.ThisAddIn.Application.ActiveSheet;

            Excel.Range actCell = Globals.ThisAddIn.Application.ActiveCell;
            if (actCell is null)
            {
                MessageBox.Show("Выберите ячейку перед запуском команды", "Ошибка");
            }

            string csvPath = GetCsvPath();
            if (File.Exists(csvPath))
            {
                var lines = GetData(csvPath).ToArray();
                var startRow = actCell.Row;
                var startCol = actCell.Column;
                var row = startRow;
                var col = startCol;

                for (int i = 0; i < lines.Length; i++)
                {
                    for (int j = 0; j < lines[i].Length; j++)
                    {
                        activeWorksheet.Cells[row, col].Value = lines[i][j];
                        col++;
                    }
                    col = startCol;
                    row++;
                }
            }
        }
    }
}
