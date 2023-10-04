using System.IO;
using ClosedXML.Excel;

namespace Sol.Core
{
    internal static class ExcelExtensions
    {
        public static byte[] GetAsByteArray(this XLWorkbook package)
        {
            using var outputMemoryStream = new MemoryStream();
            package.SaveAs(outputMemoryStream);
            return outputMemoryStream.ToArray();
        }
    }
}
