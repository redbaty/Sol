using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace Sol.Core
{
    public class JsonConverter : IConverter
    {
        public Task<byte[]> Convert(Stream excelStream, SolConverterOptions options)
        {
            using var package = new XLWorkbook(excelStream);

            var currentRow = 1;
            var worksheet = package.Worksheets.First();
            var reverseColumnMapping = ReadColumns(ref currentRow, worksheet).ToDictionary(i => i.Value, i => i.Key);
            var values = ReadValues(currentRow, worksheet, reverseColumnMapping, options);
            return Task.FromResult(WriteValues(values, options));
        }

        private static Dictionary<string, int> ReadColumns(ref int row, IXLWorksheet worksheet)
        {
            var maxColumn = worksheet.ColumnCount();
            var toReturn = new Dictionary<string, int>();

            for (var currentColumn = 1; currentColumn <= maxColumn; currentColumn++)
            {
                var cell = worksheet.Cell(row, currentColumn);
                var cellValue = cell.Value.ToString();

                if (string.IsNullOrEmpty(cellValue))
                    continue;

                toReturn.Add(cellValue, currentColumn);
            }

            row++;

            return toReturn;
        }

        private static IEnumerable<Dictionary<string, string>> ReadValues(int startingRow, IXLWorksheet worksheet, Dictionary<int, string> reverseColumnMapping, SolConverterOptions options)
        {
            var maxRow = worksheet.RowCount();

            for (var row = startingRow; row <= maxRow; row++)
            {
                var rowData = new Dictionary<string, string>();

                for (var currentColumn = 1; currentColumn <= reverseColumnMapping.Max(i => i.Key); currentColumn++)
                {
                    var cell = worksheet.Cell(row, currentColumn);
                    var cellValue = cell.Value.ToString();
                    rowData.Add(reverseColumnMapping[currentColumn], cellValue);
                }

                if (!options.IgnoreNullOnlyRows || rowData.Any(i => i.Value != null))
                    yield return rowData;
            }
        }

        private static byte[] WriteValues(IEnumerable<Dictionary<string, string>> values, SolConverterOptions options)
        {
            return JsonSerializer.SerializeToUtf8Bytes(values, new JsonSerializerOptions
            {
                WriteIndented = options.WriteFormatted
            });
        }
    }
}