namespace Sol.Core
{
    public class SolConverterOptions
    {
        public bool IgnoreNullOnlyRows { get; set; }

        public bool WriteFormatted { get; set; }
        
        public string Root { get; set; }
        
        public static SolConverterOptions Default => new()
        {
            IgnoreNullOnlyRows = true,
            WriteFormatted = true,
            Root = null
        };
    }
}