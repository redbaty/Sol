using Newtonsoft.Json.Linq;
using OfficeOpenXml;
using OfficeOpenXml.FormulaParsing.Excel.Functions.Information;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;

namespace Sol.Core
{
    public static class SolConverter
    {
        public static async Task<byte[]> ToXlsx(Stream? jsonStream, SolConverterOptions? options)
        {
            if (jsonStream is null)
            {
                throw new ArgumentNullException(nameof(jsonStream));
            }

            if (options is null)
            {
                options = SolConverterOptions.Default;
            }

            var jDoc = await JsonDocument.ParseAsync(jsonStream);
            var doc = jDoc.RootElement;
            var rawColumns = new List<string>();

            if (doc.ValueKind == JsonValueKind.Array)
            {
                rawColumns.AddRange(GetColumnsFromArray(doc).Distinct());
            }

            var columnMapping = rawColumns.Select((i, n) => new { Header = i, ColumnIndex = n + 1 }).ToDictionary(i => i.Header, i => i.ColumnIndex);

            var package = new ExcelPackage();
            var worksheet = package.Workbook.Worksheets.Add("json");
            var currentRow = 1;

            WriteHeaders(ref currentRow, worksheet, columnMapping, options);
            WriteValues(ref currentRow, worksheet, columnMapping, doc, options);
            worksheet.AutoFit();

            return package.GetAsByteArray();
        }

        public static byte[] ToJson(Stream? excelStream, SolConverterOptions? options)
        {
            if (excelStream is null)
            {
                throw new ArgumentNullException(nameof(excelStream));
            }

            if (options is null)
            {
                options = SolConverterOptions.Default;
            }

            var package = new ExcelPackage();
            package.Load(excelStream);

            var currentRow = 1;
            var worksheet = package.Workbook.Worksheets.First();
            var reverseColumnMapping = ReadColumns(ref currentRow, worksheet, options).ToDictionary(i => i.Value, i => i.Key);
            var values = ReadValues(currentRow, worksheet, reverseColumnMapping, options);
            return WriteValues(values, options);
        }

        private static Dictionary<string, int> ReadColumns(ref int row, ExcelWorksheet worksheet, SolConverterOptions options)
        {
            var maxColumn = worksheet.Dimension.Columns;
            var toReturn = new Dictionary<string, int>();

            for (int currentColumn = 1; currentColumn <= maxColumn; currentColumn++)
            {
                var cell = worksheet.Cells[row, currentColumn];
                var cellValue = cell.Value?.ToString();

                if (string.IsNullOrEmpty(cellValue))
                    continue;

                toReturn.Add(cellValue, currentColumn);
            }

            row++;

            return toReturn;
        }

        private static IEnumerable<Dictionary<string, string>> ReadValues(int startingRow, ExcelWorksheet worksheet, Dictionary<int, string> reverseColumnMapping, SolConverterOptions options)
        {
            var maxRow = worksheet.Dimension.Rows;

            for (var row = startingRow; row <= maxRow; row++)
            {
                var maxColumn = worksheet.Dimension.Columns;
                var rowData = new Dictionary<string, string>();

                for (int currentColumn = 1; currentColumn <= reverseColumnMapping.Max(i => i.Key); currentColumn++)
                {
                    var cell = worksheet.Cells[row, currentColumn];
                    var cellValue = cell.Value?.ToString();
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

        private static void WriteValues(ref int currentRow, ExcelWorksheet worksheet, Dictionary<string, int> columnMappings, JsonElement doc, SolConverterOptions options)
        {
            foreach (var node in doc.EnumerateArray())
            {
                foreach (var prop in node.EnumerateObject())
                {
                    var columnMapping = columnMappings[prop.Name];
                    var value = prop.Value.ValueKind == JsonValueKind.String ? prop.Value.GetString() : prop.Value.GetRawText();
                    worksheet.Cells[currentRow, columnMapping].Value = value;
                }

                currentRow++;
            }
        }

        private static void WriteHeaders(ref int currentRow, ExcelWorksheet worksheet, Dictionary<string, int> columnMapping, SolConverterOptions options)
        {
            foreach (var (key, value) in columnMapping)
            {
                worksheet.Cells[currentRow, value].Value = key;
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
