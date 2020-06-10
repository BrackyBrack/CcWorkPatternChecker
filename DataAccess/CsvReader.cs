using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using System.Runtime.InteropServices;

namespace DataAccess
{
    public static class CsvReader
    {

        public static DataTable GetPanels()
        {
            DataSet ds = new DataSet();
            ds.Tables.Add(ExcelToDataTable("C:\\Users\\david.bracken\\OneDrive - TUI\\Documents\\Furlough\\CC Infor\\New Panels.xlsx"));
            return ds.Tables[0];
        }

        public static DataTable GetCrewPanels()
        {
            return ExcelToDataTable("C:\\Users\\david.bracken\\OneDrive - TUI\\Documents\\Furlough\\CC Infor\\NewPatterns.xlsx");
        }

        private static DataTable ExcelToDataTable(string excelFilePath)
        {
            Microsoft.Office.Interop.Excel.Application application = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook = application.Workbooks.Open(excelFilePath, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Microsoft.Office.Interop.Excel.Worksheet worksheet = workbook.Worksheets[1];

            DataTable dt = WorksheetToDataTable(worksheet);
            workbook.Close();
            Marshal.ReleaseComObject(worksheet);
            Marshal.ReleaseComObject(workbook);
            Marshal.ReleaseComObject(application);
            return dt;
        }

        private static DataTable WorksheetToDataTable(Microsoft.Office.Interop.Excel.Worksheet worksheet)
        {
            Microsoft.Office.Interop.Excel.Range range = worksheet.UsedRange;
            object[,] valueArray = (object[,])range.Value;

            DataTable dt = new DataTable();

            //Add columns and set column names to values in row 1 of worksheet
            for (int i = 1; i < valueArray.GetUpperBound(1) + 1; i++)
            {
                string columnName = valueArray[1, i]?.ToString();

                if (string.IsNullOrWhiteSpace(columnName))
                {
                    columnName = "Column " + i.ToString();
                }

                dt.Columns.Add(columnName);
            }

            //Add rows to datatable starting at row 2 of worksheet
            for (int i = 2; i < valueArray.GetUpperBound(0) + 1; i++)
            {
                DataRow row = dt.NewRow();
                int cell = 0;
                for (int j = 1; j < dt.Columns.Count + 1; j++)
                {
                    row[cell] = valueArray[i, j];
                    cell++;
                }
                dt.Rows.Add(row);
            }

            return dt;
        }
    }
}
