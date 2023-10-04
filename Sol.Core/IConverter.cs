using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Threading.Tasks;

namespace Sol.Core
{
    public interface IConverter
    {
        Task<byte[]> Convert([NotNull]
            Stream stream, [NotNull]
            SolConverterOptions options);
    }
}