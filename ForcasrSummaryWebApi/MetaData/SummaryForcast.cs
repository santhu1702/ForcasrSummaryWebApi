﻿// <auto-generated> This file has been auto generated by EF Core Power Tools. </auto-generated>
#nullable disable
using System;
using System.Collections.Generic;

namespace ForcasrSummaryWebApi.MetaData
{
    public partial class SummaryForcast
    {
        public int SummaryForcastId { get; set; }
        public int Period { get; set; }
        public DateTime? Week { get; set; }
        public long? Year { get; set; }
        public string Category { get; set; }
        public string SubCategory { get; set; }
        public string Source { get; set; }
        public string Brand { get; set; }
        public decimal? UldollarSales { get; set; }
        public decimal? TtldollarSales { get; set; }
        public bool? IsActive { get; set; }
        public string Createdby { get; set; }
        public DateTime? CreatedDate { get; set; }
        public string UpdatedBy { get; set; }
        public DateTime? UpdatedDate { get; set; }
    }
}