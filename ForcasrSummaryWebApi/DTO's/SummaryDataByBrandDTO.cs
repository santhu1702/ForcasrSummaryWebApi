namespace ForcasrSummaryWebApi.DTO_s
{
    public class SummaryDataByBrandDTO
    {
        public string[]? subCategory { get;  set; } 
        public string[]? DataSource { get; set; }
        public string[]? rollup { get; set; }

        public string[]? Brands { get; set;}
        public int[]? Years { get; set;}

        public string? Measure { get; set;}
    }
}
