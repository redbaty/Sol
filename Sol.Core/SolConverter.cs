using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using ClosedXML.Excel;

namespace Sol.Core
{
    public static class SolConverter
    {
#nullable enable
        public static async Task<byte[]> ToXlsx(Stream? jsonStream, SolConverterOptions? options)
        {
            if (jsonStream is null)
            {
                throw new ArgumentNullException(nameof(jsonStream));
            }

            options ??= SolConverterOptions.Default;
            var jDoc = await JsonDocument.ParseAsync(jsonStream);
            var doc = jDoc.RootElement;
            var rawColumns = new List<string>();

            if (doc.ValueKind == JsonValueKind.Array)
            {
                rawColumns.AddRange(GetColumnsFromArray(doc).Distinct());
            }

            var columnMapping = rawColumns.Select((i, n) => new { Header = i, ColumnIndex = n + 1 }).ToDictionary(i => i.Header, i => i.ColumnIndex);

            using var package = new XLWorkbook();
            var worksheet = package.Worksheets.Add("json");
            var currentRow = 1;

            WriteHeaders(ref currentRow, worksheet, columnMapping);
            WriteValues(ref currentRow, worksheet, columnMapping, doc);

            using var outputMemoryStream = new MemoryStream();
            package.SaveAs(outputMemoryStream);
            return outputMemoryStream.ToArray();
        }

        public static byte[] ToJson(Stream? excelStream, SolConverterOptions? options)
        {
            if (excelStream is null)
            {
                throw new ArgumentNullException(nameof(excelStream));
            }

            options ??= SolConverterOptions.Default;

            using var package = new XLWorkbook(excelStream);
            var currentRow = 1;
            var worksheet = package.Worksheets.First();
            var reverseColumnMapping = ReadColumns(ref currentRow, worksheet).ToDictionary(i => i.Value, i => i.Key);
            var values = ReadValues(currentRow, worksheet, reverseColumnMapping, options);
            return WriteValues(values, options);
        }
#nullable restore

        private static Dictionary<string, int> ReadColumns(ref int row, IXLWorksheet worksheet)
        {
            var maxColumn = worksheet.ColumnCount();
            var toReturn = new Dictionary<string, int>();

            for (int currentColumn = 1; currentColumn <= maxColumn; currentColumn++)
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

                for (int currentColumn = 1; currentColumn <= reverseColumnMapping.Max(i => i.Key); currentColumn++)
                {
                    var cell = worksheet.Cell(row, currentColumn);
                    var cellValue = cell.Value.ToString();
                    rowData.Add(reverseColumnMapping[currentColumn], cellValue);
                }

                if(!options.IgnoreNullOnlyRows || rowData.Any(i => i.Value != null))
                    yield return rowData;
            }
        }

        private static byte[] WriteValues(IEnumerable<Dictionary<string, string>> values, SolConverterOptions options)
        {
            var json = Newtonsoft.Json.JsonConvert.SerializeObject(values, options.WriteFormatted ? Newtonsoft.Json.Formatting.Indented : Newtonsoft.Json.Formatting.None);
            return Encoding.UTF8.GetBytes(json);
        }

        private static void WriteValues(ref int currentRow, IXLWorksheet worksheet, Dictionary<string, int> columnMappings, JsonElement doc)
        {
            foreach (var node in doc.EnumerateArray())
            {
                foreach (var prop in node.EnumerateObject())
                {
                    var columnMapping = columnMappings[prop.Name];
                    var value = prop.Value.ValueKind == JsonValueKind.String ? prop.Value.GetString() : prop.Value.GetRawText();
                    worksheet.Cell(currentRow, columnMapping).Value = value;
                }

                currentRow++;
            }
        }

        private static void WriteHeaders(ref int currentRow, IXLWorksheet worksheet, Dictionary<string, int> columnMapping)
        {
            foreach (var (key, value) in columnMapping)
            {
                worksheet.Cell(currentRow, value).Value = key;
            }

            currentRow++;
        }

        private static IEnumerable<string> GetColumnsFromArray(JsonElement doc)
        {
            foreach (var node in doc.EnumerateArray())
            {
                foreach (var property in node.EnumerateObject())
                {
                    yield return property.Name;
                }
            }
        }
    }
}
