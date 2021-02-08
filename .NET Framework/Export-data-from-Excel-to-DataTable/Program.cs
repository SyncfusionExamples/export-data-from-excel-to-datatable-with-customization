using Syncfusion.XlsIO;
using System.Data;
using System.IO;


namespace ExcelToDataTable
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

                //Read data from the worksheet and exports to the data table
                DataTable customersTable = worksheet.ExportDataTable(worksheet.UsedRange, ExcelExportDataTableOptions.ColumnNames);

                //Saving the workbook as stream
                FileStream stream = new FileStream("ExportToDT.xlsx", FileMode.Create, FileAccess.ReadWrite);
                workbook.SaveAs(stream);
                stream.Dispose();
            }
        }

    }
}

