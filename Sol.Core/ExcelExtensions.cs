using OfficeOpenXml;

namespace Sol.Core
{
    internal static class ExcelExtensions
    {
        public static void AutoFit(this ExcelPackage package)
        {
            package.Workbook.AutoFit();
        }

        public static void AutoFit(this ExcelWorkbook workbook)
        {
            foreach (var worksheet in workbook.Worksheets)
            {
                worksheet.AutoFit();
            }
        }

        public static void AutoFit(this ExcelWorksheet sheet)
        {
            if (sheet.Dimension != null)
                sheet.Cells[sheet.Dimension.Address]
                    .AutoFitColumns();
        }
    }
}
