using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;  // XLSX
using NPOI.HSSF.UserModel;  // XLS
using System.IO;
using System.Text;
using System;

namespace ExcelToCsvApp
{
    public static class ExcelConverter
    {
        public static string ConvertToCsv(string filePath)
        {
            IWorkbook workbook;
            using (var fs = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                if (Path.GetExtension(filePath).ToLower() == ".xls")
                    workbook = new HSSFWorkbook(fs);  // XLS
                else
                    workbook = new XSSFWorkbook(fs);  // XLSX
            }

            var sheet = workbook.GetSheetAt(0);
            var sb = new StringBuilder();

            for (int row = sheet.FirstRowNum; row <= sheet.LastRowNum; row++)
            {
                var rowData = sheet.GetRow(row);
                if (rowData == null) continue;

                int lastCellWithValue = -1;
                for (int col = rowData.LastCellNum - 1; col >= 0; col--)
                {
                    var cell = rowData.GetCell(col);
                    if (cell != null && !string.IsNullOrWhiteSpace(cell.ToString()))
                    {
                        lastCellWithValue = col;
                        break;
                    }
                }

                if (lastCellWithValue < 0)
                {
                    continue;
                }

                var cells = new string[lastCellWithValue + 1];
                for (int col = 0; col <= lastCellWithValue; col++)
                {
                    var cell = rowData.GetCell(col);
                    cells[col] = EscapeCsvField(cell?.ToString() ?? "");
                }
                sb.AppendLine(string.Join(",", cells));
            }

            return sb.ToString();
        }

        private static string EscapeCsvField(string field)
        {
            if (field == null)
                return "";

            bool mustQuote = field.Contains(",") || field.Contains("\"") || field.Contains("\n") || field.Contains("\r");

            if (mustQuote)
            {
                var escaped = field.Replace("\"", "\"\"");
                return $"\"{escaped}\"";
            }
            else
            {
                return field;
            }
        }
    }
}
