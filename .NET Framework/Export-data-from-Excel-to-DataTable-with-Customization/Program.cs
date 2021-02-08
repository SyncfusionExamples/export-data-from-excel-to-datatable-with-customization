using Syncfusion.XlsIO;
using Syncfusion.XlsIO.Implementation;
using System.Data;
using System.IO;


namespace ExcelToDataTableCustom
{
    class Program
    {
        static void Main(string[] args)
        {
            using (ExcelEngine excelEngine = new ExcelEngine())
            {
                IApplication application = excelEngine.Excel;
                application.DefaultVersion = ExcelVersion.Xlsx;
                FileStream inputStream = new FileStream("../../Data/Sample.xlsx", FileMode.Open, FileAccess.Read);
                IWorkbook workbook = application.Workbooks.Open(inputStream);
                IWorksheet worksheet = workbook.Worksheets[0];

                //Event to choose an action while exporting data from Excel to data table.
                worksheet.ExportDataTableEvent += ExportDataTable_EventAction;

                //Read data from the worksheet and exports to the data table
                DataTable customersTable = worksheet.ExportDataTable(worksheet.UsedRange, ExcelExportDataTableOptions.ColumnNames);

                //Saving the workbook as stream
                FileStream stream = new FileStream("ExportToDT.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(stream);
                stream.Dispose();
            }

        }
        private static void ExportDataTable_EventAction(ExportDataTableEventArgs e)
        {
            if (e.ExcelValue != null && e.ExcelValue.ToString() == "Owner")
            {
                //Skips the row to export into the data table if the Excel cell value is “Owner”
                e.ExportDataTableAction = ExportDataTableActions.SkipRow;
            }
            else if (e.DataTableColumnIndex == 0 && e.ExcelRowIndex == 5 && e.ExcelColumnIndex == 1)
            {
                //Stops the export based on the condition
                e.ExportDataTableAction = ExportDataTableActions.StopExporting;
            }
            else if (e.ExcelValue != null && e.ExcelValue.ToString() == "Mexico D.F.")
            {
                //Replaces the cell value in the data table without affecting the Excel document.
                e.DataTableValue = "Mexico";
            }
        }

    }
}
