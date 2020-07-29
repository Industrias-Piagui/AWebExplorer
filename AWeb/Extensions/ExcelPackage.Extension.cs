using OfficeOpenXml;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;

namespace AWeb.Extensions
{
    public static class ExcelPackageExtension
    {
        public static byte[] GetCsv(this ExcelPackage excel, ExcelWorksheet sheet)
        {
            var maxColumnNumber = sheet.Dimension.End.Column;
            var currentRow = new List<string>(maxColumnNumber);
            var totalRowCount = sheet.Dimension.End.Row;
            var str = new StringBuilder();

            for (var i = 1; i <= totalRowCount; i++)
                str.Append(GetRowAsCsv(sheet, i, maxColumnNumber));

            return Encoding.UTF8.GetBytes(str.ToString());
        }

        public static void MergeSheets(this ExcelWorksheet sheet, ExcelWorksheet anotherSheet)
        {
            var maxColumnNumber = anotherSheet.Dimension.End.Column;
            var totalRowCount = sheet.Dimension?.End?.Row ?? 0;
            var totalRowCount2 = anotherSheet.Dimension.End.Row;

            for (var row = 1; row <= totalRowCount2; row++)
                for (var col = 1; col <= maxColumnNumber; col++)
                    sheet.Cells[row + totalRowCount, col].Value = anotherSheet.Cells[row, col]?.Value.ToString().Replace(",", "|");
        }

        private static string GetRowAsCsv(ExcelWorksheet sheet, int row, int maxColumnNumber)
        {
            var strRow = new StringBuilder();
            for (var i = 1; i <= maxColumnNumber; i++)
            {
                strRow.Append(GetCellAsCsv(sheet, row, i));
                strRow.Append(i == maxColumnNumber ? "\r\n" : ",");
            }
            return strRow.ToString();
        }

        private static string GetCellAsCsv(ExcelWorksheet sheet, int row, int col)
        {
            var cell = sheet.Cells[row, col];
            var str = string.IsNullOrEmpty(cell?.Value?.ToString()) ? "" : cell.Value.ToString();
            str = str.Replace(",", "");
            str = RemoveAccentMark(str);
            if (str.Contains('\n') || str.Contains('\"'))
                return $"\"{str}\"";

            return str;
        }

        private static string RemoveAccentMark(string str)
        {
            var tmpBytes = Encoding.GetEncoding("ISO-8859-8").GetBytes(str);
            return Encoding.UTF8.GetString(tmpBytes);
        }
    }
}
