using System;
namespace ForcasrSummaryWebApi.DTOs
{
    public class CombinedDataDTO
    {
        public List<string>? SubCategories { get; set; }
        public List<string>? Sources { get; set; }
        public List<string>? Brands { get; set; }
        public List<long>? Years { get; set; }
    }
}

