using Sol.Core;
using System.Threading.Tasks;
using CliFx;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Serilog;
using Serilog.Extensions.Logging;

namespace Sol
{
    internal static class Program
    {
        private static async Task<int> Main(string[] args)
        {
            var loggerConfiguration = new LoggerConfiguration()
                .Enrich.FromLogContext()
                .MinimumLevel.Debug()
                .WriteTo.Console();
            
            Log.Logger = loggerConfiguration.CreateLogger();
            
            var serviceCollection = new ServiceCollection();
            serviceCollection.AddTransient<JsonConverter>();
            serviceCollection.AddTransient<XlsxConverter>();
            serviceCollection.AddTransient<ConvertCommand>();
            serviceCollection.AddSingleton<ILoggerFactory>(_ => new SerilogLoggerFactory());
            serviceCollection.AddLogging(options => options.AddSerilog());
            var serviceProvider = serviceCollection.BuildServiceProvider();

            return await new CliApplicationBuilder()
                .AddCommand<ConvertCommand>()
                .UseTypeActivator(serviceProvider.GetService)
                .Build()
                .RunAsync(args);
        }
    }
}
