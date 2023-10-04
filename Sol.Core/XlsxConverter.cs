using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using ClosedXML.Excel;
using JsonExtensions.Reading;
using Microsoft.Extensions.Options;

namespace Sol.Core
{
    public class XlsxConverter : IConverter
    {
        public XlsxConverter(IOptions<XlsxOptions> options)
        {
            Options = options;
        }

        private IOptions<XlsxOptions> Options { get; }

        private record JsonRowValue(string Content, string Path)
        {
            public int PathLength => Path?.Count(i => i == '.') ?? 0;
        }

        private record Row(JsonRowValue[] Values);

        private static IEnumerable<JsonRowValue> ReadElementValues(JsonElement rootElement, string basePath)
        {
            foreach (var jsonProperty in rootElement.EnumerateObject())
            {
                if (jsonProperty.Value.ValueKind != JsonValueKind.Array && jsonProperty.Value.ValueKind != JsonValueKind.Object)
                {
                    yield return new JsonRowValue(jsonProperty.Value.GetRawText(), $"{basePath}{jsonProperty.Name}");
                }
                else
                {
                    if (jsonProperty.Value.ValueKind == JsonValueKind.Object)
                    {
                        foreach (var valor in ReadElementValues(jsonProperty.Value, $"{jsonProperty.Name}."))
                        {
                            yield return valor;
                        }
                    }
                }
            }
        }

        private IEnumerable<Row> ReadElementAsRow(JsonElement rootElement)
        {
            switch (rootElement.ValueKind)
            {
                case JsonValueKind.Object:
                    if (Options.Value.PropertiesAsRow)
                        foreach (var jsonElement in rootElement.EnumerateObject())
                            yield return new Row(ReadElementValues(jsonElement.Value, string.Empty).Concat(new[] { new JsonRowValue(jsonElement.Name, "property_name") }).Reverse().ToArray());
                    else
                        yield return new Row(ReadElementValues(rootElement, string.Empty).ToArray());
                    break;
                case JsonValueKind.Array:
                {
                    foreach (var jsonElement in rootElement.EnumerateArray())
                        yield return new Row(ReadElementValues(jsonElement, string.Empty).ToArray());
                    break;
                }
                default:
                    throw new NotImplementedException();
            }
        }

        public async Task<byte[]> Convert(Stream jsonStream, SolConverterOptions options)
        {
            var jDoc = await JsonDocument.ParseAsync(jsonStream);
            var docRoot = jDoc.RootElement;

            if (!string.IsNullOrEmpty(options.Root))
            {
                docRoot = docRoot.GetPropertyByPath(options.Root);
            }

            var linhas = ReadElementAsRow(docRoot).ToArray();
            var headers = linhas
                .SelectMany(i => i.Values.Select(o => new { o.Path, o.PathLength }))
                .Distinct()
                .OrderBy(i => i.PathLength)
                .Select((i, index) => new { i.Path, Column = index })
                .ToDictionary(i => i.Path, i => i.Column);
            using var package = new XLWorkbook();
            var worksheet = package.Worksheets.Add("json");
            var currentRow = 1;

            {
                var headerIndex = 1;

                foreach (var headerName in headers.Keys)
                {
                    worksheet.Cell(currentRow, headerIndex).Value = headerName;
                    headerIndex++;
                }
            }

            currentRow++;
            foreach (var linha in linhas)
            {
                foreach (var valor in linha.Values)
                {
                    var header = headers[valor.Path];
                    worksheet.Cell(currentRow, header + 1).Value = valor.Content;
                }

                currentRow++;
            }
            
            return package.GetAsByteArray();
        }
    }
}