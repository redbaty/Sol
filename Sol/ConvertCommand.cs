using System;
using System.IO;
using System.Threading.Tasks;
using CliFx;
using CliFx.Attributes;
using CliFx.Exceptions;
using CliFx.Infrastructure;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Sol.Core;

namespace Sol
{
    [Command]
    public class ConvertCommand : ICommand
    {
        public ConvertCommand(IServiceProvider serviceProvider, ILogger<ConvertCommand> logger, IOptions<XlsxOptions> xlsxOptions)
        {
            ServiceProvider = serviceProvider;
            Logger = logger;
            XlsxOptions = xlsxOptions;
        }

        [CommandParameter(0, Description = "Full file path.")]
        public string FilePath { get; set; }

        [CommandOption("conversion-type", 'c', Description = "Force conversion type.")]
        public ConversionType? Type { get; set; }

        [CommandOption("write-formatted", 'f', Description = "Write formatted json.")]
        public bool WriteFormatted { get; set; } = false;

        [CommandOption("ignore-null-rows", 'n', Description = "Ignore null rows when reading from excel.")]
        public bool IgnoreNullOnlyRows { get; set; } = true;
        
        [CommandOption("output-file", 'o')]
        public string OutputFile { get; set; }
        
        [CommandOption("root", 'r')]
        public string Root { get; set; }
        
        [CommandOption("property-as-row")]
        public bool PropertyAsRow { get; set; }

        private IServiceProvider ServiceProvider { get; }

        private ILogger<ConvertCommand> Logger { get; }
        
        private IOptions<XlsxOptions> XlsxOptions { get; }

        public async ValueTask ExecuteAsync(IConsole console)
        {
            var fileInfo = string.IsNullOrEmpty(FilePath) ? null : new FileInfo(FilePath);

            if (fileInfo is not {Exists: true}) throw new CommandException("File doesn't exist.");
            XlsxOptions.Value.PropertiesAsRow = PropertyAsRow;
            
            var options = BuildOptions();
            await using var fileStream = fileInfo.OpenRead();
            var conversionType = Type ?? GetConversionType(fileInfo.Extension);
            Logger.LogInformation("Conversion type defined as: {@ConversionType}", conversionType);

            var outputFileName = OutputFile ?? ChangeFileExtension(fileInfo, GetNewFileExtension(conversionType));
            var result = await Convert(conversionType, fileStream, options);
            await File.WriteAllBytesAsync(outputFileName, result);
            Logger.LogInformation("\"{@InputFile}\" was successfully converted to \"{@OutputFile}\"", fileInfo.FullName, outputFileName);
        }

        private SolConverterOptions BuildOptions()
        {
            return new SolConverterOptions
            {
                WriteFormatted = WriteFormatted,
                IgnoreNullOnlyRows = IgnoreNullOnlyRows,
                Root = Root
            };
        }

        private IConverter GetConverter(ConversionType conversionType)
        {
            return conversionType switch
            {
                ConversionType.XlsxToJson => ServiceProvider.GetService<JsonConverter>(),
                ConversionType.JsonToXls => ServiceProvider.GetService<XlsxConverter>(),
                _ => throw new ArgumentOutOfRangeException(nameof(conversionType))
            };
        }

        private static string GetNewFileExtension(ConversionType conversionType)
        {
            return conversionType switch
            {
                ConversionType.XlsxToJson => ".json",
                ConversionType.JsonToXls => ".xlsx",
                _ => throw new ArgumentOutOfRangeException(nameof(conversionType))
            };
        }

        private static ConversionType GetConversionType(string extension)
        {
            return extension switch
            {
                ".json" => ConversionType.JsonToXls,
                ".xls" => ConversionType.XlsxToJson,
                ".xlsx" => ConversionType.XlsxToJson,
                _ => throw new CommandException("Can't determine conversion type, please specify one using the \"-c/--conversion-type\" flag.", 2)
            };
        }

        private Task<byte[]> Convert(ConversionType conversionType, Stream fileStream, SolConverterOptions options = null)
        {
            return GetConverter(conversionType).Convert(fileStream, options);
        }

        private static string ChangeFileExtension(FileInfo file, string extension)
        {
            var fileName = $"{Path.GetFileNameWithoutExtension(file.FullName)}{extension}";
            return string.IsNullOrEmpty(file.DirectoryName) ? fileName : Path.Combine(file.DirectoryName, fileName);
        }
    }
}