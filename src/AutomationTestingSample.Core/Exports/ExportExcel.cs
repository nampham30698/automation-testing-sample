using OfficeOpenXml;

namespace AutomationTestingSample.Core.Exports
{
    public class ExportExcel
    {
        public void Export()
        {
            ExcelPackage excel = new ExcelPackage();

            // name of the sheet 
            var workSheet = excel.Workbook.Worksheets.Add("Sheet1");
        }
    }
}
