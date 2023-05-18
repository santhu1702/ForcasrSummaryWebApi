using System;
using System.Data;
using ForcasrSummaryWebApi.CommonMethods;
using ForcasrSummaryWebApi.DTO_s;
using ForcasrSummaryWebApi.DTOs;
using ForcasrSummaryWebApi.MetaData;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;

namespace ForcasrSummaryWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        public static string _headers = "Total,C1,C2,C3,C4,C5,C6,C7,C8,C9,C10,C11,C12,C13";
        public static string _quaterHeaders = "FY MAT,H1,H2, Q1,Q2,Q3,Q4";
        private readonly budgetForecastContext? _context;
        public ExcelController(budgetForecastContext? context)
        {
            _context = context;
        }



        [HttpGet("DropDownsData")]
        public async Task<ActionResult<CombinedDataDTO>> GetDropDownsData()
        {
            try
            {
                var subCategories = await _context.SummaryForcast
                    .Select(x => x.SubCategory)
                    .Distinct()
                    .ToListAsync();

                var sources = await _context.SummaryForcast
                    .Select(x => x.Source)
                    .Distinct()
                    .ToListAsync();

                var brands = await _context.SummaryForcast
                    .Select(x => x.Brand)
                    .Distinct()
                    .ToListAsync();

                var years = await _context.SummaryForcast
                    .Select(x => x.Year)
                    .Distinct()
                    .ToListAsync();

                var combinedData = new CombinedDataDTO
                {
                    SubCategories = subCategories,
                    Sources = sources,
                    Brands = brands,
                    Years = years.Select(x => x.Value).ToList()
                };

                return Ok(combinedData);
            }
            catch (Exception ex)
            {
                // Handle the exception, log it, and return an error response
                return StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }



        [HttpGet("SummaryData")]
        public async Task<ActionResult<IEnumerable<object>>> GetSummaryData()
        {
            var categoryColumnNo = 2;
            var headers = _headers.Replace("SubCategory", "xyz").Split(',');
            var result = new string[headers.Length + 1];
            result[0] = "Body Cleansing";

            for (var i = 1; i < result.Length; i++)
            {
                result[i] = headers[i - 1];
            }

            var summaryDataArray = new List<string[]> { result };
            var dataList = await _context.GetProcedures().USP_GetSummaryDataAsync();
            var dataByYear = dataList.GroupBy(d => d.Year).OrderBy(g => g.Key);
            List<Dictionary<string, int>> mergeArray = new List<Dictionary<string, int>>();
            var subcname = "BAR SOAP";
            var getDataByBrand = await _context.GetProcedures().USP_GetSummaryDataByBrandAsync("BABY CLEANSING,BAR SOAP", "IRI", "BODY CLEANSING", "Brand 1,Brand 2", "2020,2021");

            foreach (var yearData in dataByYear)
            {
                foreach (var data in yearData)
                {
                    var nestedvalues = new[] { data.Year.ToString(), data.SalesType, data._1.ToString(), data._2.ToString(), data._3.ToString(), data._4.ToString(), data._5.ToString(), data._6.ToString(), data._7.ToString(), data._8.ToString(), data._9.ToString(), data._10.ToString(), data._11.ToString(), data._12.ToString(), data._13.ToString() };
                    summaryDataArray.Add(nestedvalues);
                }

                // var additionalValues= Array.Empty<string>();
                if (dataByYear.Last().Key == yearData.Key)
                {
                    //need to implement
                    var additionalValues = new[] { yearData.First().Year.ToString(), "$ % Chg", $"=C{categoryColumnNo + 1}/ C{categoryColumnNo - 3}-1", $"=D{categoryColumnNo + 1}/ D{categoryColumnNo - 3}-1", $"=E{categoryColumnNo + 1}/ E{categoryColumnNo - 3}-1", $"=F{categoryColumnNo + 1}/ F{categoryColumnNo - 3}-1", $"=G{categoryColumnNo + 1}/ G{categoryColumnNo - 3}-1", $"=H{categoryColumnNo + 1}/ H{categoryColumnNo - 3}-1", $"=I{categoryColumnNo + 1}/ I{categoryColumnNo - 3}-1", $"=J{categoryColumnNo + 1}/ J{categoryColumnNo - 3}-1", $"=K{categoryColumnNo + 1}/ K{categoryColumnNo - 3}-1", $"=L{categoryColumnNo + 1}/ L{categoryColumnNo - 3}-1", $"=M{categoryColumnNo + 1}/ M{categoryColumnNo - 3}-1", $"=N{categoryColumnNo + 1}/ N{categoryColumnNo - 3}-1", $"=O{categoryColumnNo + 1}/ O{categoryColumnNo - 3}-1" };
                    var additionalValues1 = new[] { yearData.First().Year.ToString(), "UL $ Share", $"=C{categoryColumnNo + 1}/ C{categoryColumnNo}*100", $"=D{categoryColumnNo + 1}/ D{categoryColumnNo}*100", $"=E{categoryColumnNo + 1}/ E{categoryColumnNo}*100", $"=F{categoryColumnNo + 1}/ F{categoryColumnNo}*100", $"=G{categoryColumnNo + 1}/ G{categoryColumnNo}*100", $"=H{categoryColumnNo + 1}/ H{categoryColumnNo}*100", $"=I{categoryColumnNo + 1}/ I{categoryColumnNo}*100", $"=J{categoryColumnNo + 1}/ J{categoryColumnNo}*100", $"=K{categoryColumnNo + 1}/ K{categoryColumnNo}*100", $"=L{categoryColumnNo + 1}/ L{categoryColumnNo}*100", $"=M{categoryColumnNo + 1}/ M{categoryColumnNo}*100", $"=N{categoryColumnNo + 1}/ N{categoryColumnNo}*100", $"=O{categoryColumnNo + 1}/ O{categoryColumnNo}*100" };
                    var additionalValues2 = new[] { yearData.First().Year.ToString(), "UL BPS Chg", $"=(C{categoryColumnNo + 3}- C{categoryColumnNo - 2})*100", $"=(D{categoryColumnNo + 3}- D{categoryColumnNo - 2})*100", $"=(E{categoryColumnNo + 3}- E{categoryColumnNo - 2})*100", $"=(F{categoryColumnNo + 3}- F{categoryColumnNo - 2})*100", $"=(G{categoryColumnNo + 3}- G{categoryColumnNo - 2})*100", $"=(H{categoryColumnNo + 3}- H{categoryColumnNo - 2})*100", $"=(I{categoryColumnNo + 3}- I{categoryColumnNo - 2})*100", $"=(J{categoryColumnNo + 3}- J{categoryColumnNo - 2})*100", $"=(K{categoryColumnNo + 3}- K{categoryColumnNo - 2})*100", $"=(L{categoryColumnNo + 3}- L{categoryColumnNo - 2})*100", $"=(M{categoryColumnNo + 3}- M{categoryColumnNo - 2})*100", $"=(N{categoryColumnNo + 3}- N{categoryColumnNo - 2})*100", $"=(O{categoryColumnNo + 3}- 0{categoryColumnNo - 2})*100" };
                    var additionalValues3 = new[] { yearData.First().Year.ToString(), "Proj Sales UL", $"=C{categoryColumnNo + 1}", $"=D{categoryColumnNo + 1}", $"=E{categoryColumnNo + 1}", $"=F{categoryColumnNo + 1}", $"=G{categoryColumnNo + 1}", $"=H{categoryColumnNo + 1}", $"=I{categoryColumnNo + 1}", $"=J{categoryColumnNo + 1}", $"=K{categoryColumnNo + 1}", $"=L{categoryColumnNo + 1}", $"=M{categoryColumnNo + 1}", $"=N{categoryColumnNo + 1}", $"=O{categoryColumnNo + 1}" };
                    var additionalValues4 = new[] { yearData.First().Year.ToString(), "Proj $ % Chg", $"=C{categoryColumnNo + 5}/ C{categoryColumnNo - 3}-1", $"=D{categoryColumnNo + 5}/ D{categoryColumnNo - 3}-1", $"=E{categoryColumnNo + 5}/ E{categoryColumnNo - 3}-1", $"=F{categoryColumnNo + 5}/ F{categoryColumnNo - 3}-1", $"=G{categoryColumnNo + 5}/ G{categoryColumnNo - 3}-1", $"=H{categoryColumnNo + 5}/ H{categoryColumnNo - 3}-1", $"=I{categoryColumnNo + 5}/ I{categoryColumnNo - 3}-1", $"=J{categoryColumnNo + 5}/ J{categoryColumnNo - 3}-1", $"=K{categoryColumnNo + 5}/ K{categoryColumnNo - 3}-1", $"=L{categoryColumnNo + 5}/ L{categoryColumnNo - 3}-1", $"=M{categoryColumnNo + 5}/ M{categoryColumnNo - 3}-1", $"=N{categoryColumnNo + 5}/ N{categoryColumnNo - 3}-1", $"=O{categoryColumnNo + 5}/ {categoryColumnNo - 3}-1" };
                    var additionalValues5 = new[] { yearData.First().Year.ToString(), "Proj UL $ Share", $"=C{categoryColumnNo + 5}/ C{categoryColumnNo}*100", $"=D{categoryColumnNo + 5}/ D{categoryColumnNo}*100", $"=E{categoryColumnNo + 5}/ E{categoryColumnNo}*100", $"=F{categoryColumnNo + 5}/ F{categoryColumnNo}*100", $"=G{categoryColumnNo + 5}/ G{categoryColumnNo}*100", $"=H{categoryColumnNo + 5}/ H{categoryColumnNo}*100", $"=I{categoryColumnNo + 5}/ I{categoryColumnNo}*100", $"=J{categoryColumnNo + 5}/ J{categoryColumnNo}*100", $"=K{categoryColumnNo + 5}/ K{categoryColumnNo}*100", $"=L{categoryColumnNo + 5}/ L{categoryColumnNo}*100", $"=M{categoryColumnNo + 5}/ M{categoryColumnNo}*100", $"=N{categoryColumnNo + 5}/ N{categoryColumnNo}*100", $"=O{categoryColumnNo + 5}/ O{categoryColumnNo}*100" };
                    var additionalValues6 = new[] { yearData.First().Year.ToString(), "Proj BPS Chg", $"=(C{categoryColumnNo + 7} - C{categoryColumnNo - 2})*100", $"=(D{categoryColumnNo + 7} - D{categoryColumnNo - 2})*100", $"=(E{categoryColumnNo + 7} - E{categoryColumnNo - 2})*100", $"=(F{categoryColumnNo + 7} - F{categoryColumnNo - 2})*100", $"=(G{categoryColumnNo + 7} - G{categoryColumnNo - 2})*100", $"=(H{categoryColumnNo + 7} - H{categoryColumnNo - 2})*100", $"=(I{categoryColumnNo + 7} - I{categoryColumnNo - 2})*100", $"=(J{categoryColumnNo + 7} - J{categoryColumnNo - 2})*100", $"=(K{categoryColumnNo + 7} - K{categoryColumnNo - 2})*100", $"=(L{categoryColumnNo + 7} - L{categoryColumnNo - 2})*100", $"=(M{categoryColumnNo + 7} - M{categoryColumnNo - 2})*100", $"=(N{categoryColumnNo + 7} - N{categoryColumnNo - 2})*100", $"=(O{categoryColumnNo + 7} - O{categoryColumnNo - 2})*100" };
                    var additionalValues7 = new[] { yearData.First().Year.ToString(), "Proj Bps vs Forecast", "0.0", "0.0", $"=(E{categoryColumnNo + 8} - E{categoryColumnNo + 4})", $"=(F{categoryColumnNo + 8} - F{categoryColumnNo + 4})", $"=(G{categoryColumnNo + 8} - G{categoryColumnNo + 4})", $"=(H{categoryColumnNo + 8} - H{categoryColumnNo + 4})", $"=(I{categoryColumnNo + 8} - I{categoryColumnNo + 4})", $"=(J{categoryColumnNo + 8} - J{categoryColumnNo + 4})", $"=(K{categoryColumnNo + 8} - K{categoryColumnNo + 4})", $"=(L{categoryColumnNo + 8} - L{categoryColumnNo + 4})", $"=(M{categoryColumnNo + 8} - M{categoryColumnNo + 4})", $"=(N{categoryColumnNo + 8} - N{categoryColumnNo + 4})", $"=(O{categoryColumnNo + 8} - O{categoryColumnNo + 4})" };
                    summaryDataArray.Add(additionalValues);
                    summaryDataArray.Add(additionalValues1);
                    summaryDataArray.Add(additionalValues2);
                    summaryDataArray.Add(additionalValues3);
                    summaryDataArray.Add(additionalValues4);
                    summaryDataArray.Add(additionalValues5);
                    summaryDataArray.Add(additionalValues6);
                    summaryDataArray.Add(additionalValues7);
                    Dictionary<string, int> mergeList1 = new Dictionary<string, int>
                    {
                        { "row", categoryColumnNo -1 },
                        { "col", 0 },
                        { "rowspan",10 },
                        { "colspan", 1 }
                    };

                    mergeArray.Add(mergeList1);

                }
                else
                {
                    var additionalValues = new[] { yearData.First().Year.ToString(), "UL $ Share", $"=C{categoryColumnNo + 1}/ C{categoryColumnNo}*100", $"=D{categoryColumnNo + 1}/ D{categoryColumnNo}*100", $"=E{categoryColumnNo + 1}/ E{categoryColumnNo}*100", $"=F{categoryColumnNo + 1}/ F{categoryColumnNo}*100", $"=G{categoryColumnNo + 1}/ G{categoryColumnNo}*100", $"=H{categoryColumnNo + 1}/ H{categoryColumnNo}*100", $"=I{categoryColumnNo + 1}/ I{categoryColumnNo}*100", $"=J{categoryColumnNo + 1}/ J{categoryColumnNo}*100", $"=K{categoryColumnNo + 1}/ K{categoryColumnNo}*100", $"=L{categoryColumnNo + 1}/ L{categoryColumnNo}*100", $"=M{categoryColumnNo + 1}/ M{categoryColumnNo}*100", $"=N{categoryColumnNo + 1}/ N{categoryColumnNo}*100", $"=O{categoryColumnNo + 1}/ O{categoryColumnNo}*100" };
                    summaryDataArray.Add(additionalValues);

                }

                //}
                var nullarray = new[] { "" };
                summaryDataArray.Add(nullarray);

                Dictionary<string, int> mergeList = new Dictionary<string, int>
                {
                    { "row", categoryColumnNo - 1 },
                    { "col", 0 },
                    { "rowspan", 3 },
                    { "colspan", 1 }
                };

                mergeArray.Add(mergeList);
                 // clsCommonMethods.bindingQuatersData(dataByYear, categoryColumnNo);
                clsCommonMethods.getsubbrand(getDataByBrand, _headers, _quaterHeaders, summaryDataArray , subcname);
                categoryColumnNo += 4;

            }
            
            var output = new
            {
                data = summaryDataArray,
                summaryDataArray = mergeArray
            };

            return Ok(output);
        }



        [HttpGet("SummaryDataByFilters")]
        public async Task<ActionResult<USP_GetSummaryDataByBrandResult>> getSummaryDataByBrand(SummaryDataByBrandDTO summaryData)
        {
            return Ok(await _context.GetProcedures().USP_GetSummaryDataByBrandAsync(string.Join(",", summaryData.subCatagoery),
                                                                                    string.Join(",", summaryData.source),
                                                                                    string.Join(",", summaryData.category),
                                                                                    string.Join(",", summaryData.brands),
                                                                                    string.Join(",", summaryData.years)));
        }


    }
}
