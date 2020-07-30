namespace Sol.Core
{
    public class SolConverterOptions
    {
        public static SolConverterOptions Default = new SolConverterOptions {IgnoreNullOnlyRows = true};

        public bool IgnoreNullOnlyRows { get; set; }

        public bool WriteFormatted { get; set; }
    }
}
