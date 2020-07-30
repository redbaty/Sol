using Sol.Core;
using System;
using System.IO;
using System.Threading.Tasks;

namespace Sol
{
    internal static class Program
    {
        private static async Task Main(string[] args)
        {
            var file = new FileInfo(args.Length == 0 ? ReadFileName() : args[0]);

            Console.WriteLine("Input file path: {0}", file.FullName);

            if (!file.Exists)
            {
                Console.WriteLine("File \"{0}\" is a lie", file.FullName);
                WaitForExitIfNecessary(args);
                return;
            }

            await using var fileStream = file.OpenRead();
            var conversionType = GetConversionType(file.Extension);

            Console.WriteLine("Conversion type is: {0}", conversionType);

            var result = await Convert(conversionType, fileStream);
            var outputFileName = ChangeFileExtension(file, GetNewFileExtension(conversionType));
            await File.WriteAllBytesAsync(outputFileName, result);

            Console.WriteLine("Successfully converted file. Saved as: {0}", outputFileName);
            WaitForExitIfNecessary(args);
        }

        private static void WaitForExitIfNecessary(string[] args)
        {
            if (args.Length == 0)
            {
                Console.WriteLine("Press any key to exit.");
                Console.ReadKey();
            }
        }

        private static string ReadFileName()
        {
            Console.WriteLine("No arguments provided, please enter a file path.");
            return Console.ReadLine();
        }

        private static string GetNewFileExtension(ConversionType conversionType)
        {
            return conversionType switch
            {
                ConversionType.XlsxToJson => ".json",
                ConversionType.JsonToXls => ".xlsx",
                _ => throw new ArgumentOutOfRangeException(nameof(conversionType)),
            };
        }

        private static ConversionType GetConversionType(string extension)
        {
            switch (extension)
            {
                case ".json":
                    return ConversionType.JsonToXls;
                case ".xls":
                case ".xlsx":
                    return ConversionType.XlsxToJson;
                default:
                    throw new ArgumentOutOfRangeException(nameof(extension));
            }
        }

        private static async Task<byte[]> Convert(ConversionType conversionType, Stream fileStream)
        {
            switch (conversionType)
            {
                case ConversionType.XlsxToJson:
                    return SolConverter.ToJson(fileStream);
                case ConversionType.JsonToXls:
                    return await SolConverter.ToXlsx(fileStream);
                default:
                    throw new ArgumentOutOfRangeException(nameof(conversionType));
            }
        }

        private static string ChangeFileExtension(FileInfo file, string extension)
            => Path.Combine(file.DirectoryName, $"{Path.GetFileNameWithoutExtension(file.FullName)}{extension}");
    }
}
