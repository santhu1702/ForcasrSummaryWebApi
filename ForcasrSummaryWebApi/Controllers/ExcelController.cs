using System.Data;
using ClosedXML.Excel;
using ForcasrSummaryWebApi.DTOs;
using ForcasrSummaryWebApi.MetaData;
using Microsoft.AspNetCore.Mvc;
using Microsoft.EntityFrameworkCore;
using Newtonsoft.Json;

namespace ForcasrSummaryWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
        public static string _headers = "Brand,C1,C2,C3,C4,C5,C6,C7,C8,C9,C10,C11,C12,C13";
        public static string _quarterMATHeaders = "Q1,Q2,Q3,Q4";
        public static string _halfYearMATHeaders = "H1,H2";
        public static string _fullYearMATHeaders = "FY MAT";
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
                var subCategories = await _context.Per9
                    .Select(x => x.Subcategory)
                    .Distinct()
                    .ToListAsync();

                var sources = await _context.Per9
                    .Select(x => x.Source)
                    .Distinct()
                    .ToListAsync();

                var brands = await _context.Per9
                    .Select(x => x.Brand)
                    .Distinct()
                    .ToListAsync();

                var years = await _context.Per9
                    .Select(x => x.Year)
                    .Distinct()
                    .ToListAsync();

                var combinedData = new CombinedDataDTO
                {
                    SubCategories = subCategories,
                    Sources = sources,
                    Brands = brands,
                    Years = years.ConvertAll(long.Parse)
                };

                return Ok(combinedData);
            }
            catch (Exception ex)
            {
                // Handle the exception, log it, and return an error response
                return StatusCode(StatusCodes.Status500InternalServerError, ex.Message);
            }
        }

        #region getData 
        [HttpGet("getData")]
        public async Task<ActionResult<IEnumerable<dynamic>>> GetData()
        {
            try
            {
                var summaryDataList = new List<List<string>>();
                var data = await _context.GetProcedures().usp_GetDimFactDataAsync();

                summaryDataList.AddRange(data.Select(row => new List<string>
                                        {
                                            row.Column01,
                                            row.Column04,
                                            row.Column06,
                                            row.Column11,
                                            row.Column12,
                                            row.Column13,
                                            row.Column14,
                                            row.Column15,
                                            row.Column16,
                                            row.Column17,
                                            row.Column18,
                                            row.Column19,
                                            row.Column20,
                                            row.Column21,
                                            row.Column22,
                                            row.Column23,
                                            row.DimFactID.ToString()
                                        }));

                return summaryDataList;
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
        #endregion
        #region
        [HttpPost("upload")]
        public async Task<IActionResult> UploadData([FromBody] object data)
        {
            try
            {
                // Process the received data here
                // You can access the data in the 'data' variable and perform any necessary operations
                var json = data.ToString();

                var msg = await _context.GetProcedures().uploadDataAsync(json);
                // Return a success response if the data was processed successfully
                return Ok(new { message = msg });
            }
            catch (Exception ex)
            {
                // Return an error response if there was an exception or processing error
                return BadRequest(new { error = ex.Message });
            }
        }

        #endregion
        #region SummaryData

        //[HttpGet("SummaryData")]
        //public async Task<ActionResult<IEnumerable<SummaryDataDTO>>> GetSummaryData()
        //{
        //    var categoryColumnNo = 2;
        //    var headers = _headers.Replace("SubCategory", "xyz").Split(',');
        //    var result = new string[headers.Length + 1];
        //    result[0] = "Body Cleansing";

        //    for (var i = 1; i < result.Length; i++)
        //    {
        //        result[i] = headers[i - 1];
        //    }

        //    var summaryDataArray = new List<string[]> { result };
        //    var dataList = await _context.GetProcedures().USP_GetSummaryDataAsync();
        //    var dataByYear = dataList.GroupBy(d => d.Year).OrderBy(g => g.Key);
        //    List<Dictionary<string, int>> mergeArray = new List<Dictionary<string, int>>();
        //    var subcname = "BAR SOAP";
        //    var getDataByBrand = await _context.GetProcedures().USP_GetSummaryDataByBrandAsync("BABY CLEANSING,BAR SOAP", "IRI", "Brand 1,Brand 2", "2020,2021");

        //    foreach (var yearData in dataByYear)
        //    {
        //        foreach (var data in yearData)
        //        {
        //            var nestedvalues = new[] { data.Year.ToString(), data.SalesType, data._1.ToString(), data._2.ToString(), data._3.ToString(), data._4.ToString(), data._5.ToString(), data._6.ToString(), data._7.ToString(), data._8.ToString(), data._9.ToString(), data._10.ToString(), data._11.ToString(), data._12.ToString(), data._13.ToString() };
        //            summaryDataArray.Add(nestedvalues);
        //        }

        //        // var additionalValues= Array.Empty<string>();
        //        if (dataByYear.Last().Key == yearData.Key)
        //        {
        //            //need to implement
        //            var additionalValues = new[] { yearData.First().Year.ToString(), "$ % Chg", $"=C{categoryColumnNo + 1}/ C{categoryColumnNo - 3}-1", $"=D{categoryColumnNo + 1}/ D{categoryColumnNo - 3}-1", $"=E{categoryColumnNo + 1}/ E{categoryColumnNo - 3}-1", $"=F{categoryColumnNo + 1}/ F{categoryColumnNo - 3}-1", $"=G{categoryColumnNo + 1}/ G{categoryColumnNo - 3}-1", $"=H{categoryColumnNo + 1}/ H{categoryColumnNo - 3}-1", $"=I{categoryColumnNo + 1}/ I{categoryColumnNo - 3}-1", $"=J{categoryColumnNo + 1}/ J{categoryColumnNo - 3}-1", $"=K{categoryColumnNo + 1}/ K{categoryColumnNo - 3}-1", $"=L{categoryColumnNo + 1}/ L{categoryColumnNo - 3}-1", $"=M{categoryColumnNo + 1}/ M{categoryColumnNo - 3}-1", $"=N{categoryColumnNo + 1}/ N{categoryColumnNo - 3}-1", $"=O{categoryColumnNo + 1}/ O{categoryColumnNo - 3}-1" };
        //            var additionalValues1 = new[] { yearData.First().Year.ToString(), "UL $ Share", $"=C{categoryColumnNo + 1}/ C{categoryColumnNo}*100", $"=D{categoryColumnNo + 1}/ D{categoryColumnNo}*100", $"=E{categoryColumnNo + 1}/ E{categoryColumnNo}*100", $"=F{categoryColumnNo + 1}/ F{categoryColumnNo}*100", $"=G{categoryColumnNo + 1}/ G{categoryColumnNo}*100", $"=H{categoryColumnNo + 1}/ H{categoryColumnNo}*100", $"=I{categoryColumnNo + 1}/ I{categoryColumnNo}*100", $"=J{categoryColumnNo + 1}/ J{categoryColumnNo}*100", $"=K{categoryColumnNo + 1}/ K{categoryColumnNo}*100", $"=L{categoryColumnNo + 1}/ L{categoryColumnNo}*100", $"=M{categoryColumnNo + 1}/ M{categoryColumnNo}*100", $"=N{categoryColumnNo + 1}/ N{categoryColumnNo}*100", $"=O{categoryColumnNo + 1}/ O{categoryColumnNo}*100" };
        //            var additionalValues2 = new[] { yearData.First().Year.ToString(), "UL BPS Chg", $"=(C{categoryColumnNo + 3}- C{categoryColumnNo - 2})*100", $"=(D{categoryColumnNo + 3}- D{categoryColumnNo - 2})*100", $"=(E{categoryColumnNo + 3}- E{categoryColumnNo - 2})*100", $"=(F{categoryColumnNo + 3}- F{categoryColumnNo - 2})*100", $"=(G{categoryColumnNo + 3}- G{categoryColumnNo - 2})*100", $"=(H{categoryColumnNo + 3}- H{categoryColumnNo - 2})*100", $"=(I{categoryColumnNo + 3}- I{categoryColumnNo - 2})*100", $"=(J{categoryColumnNo + 3}- J{categoryColumnNo - 2})*100", $"=(K{categoryColumnNo + 3}- K{categoryColumnNo - 2})*100", $"=(L{categoryColumnNo + 3}- L{categoryColumnNo - 2})*100", $"=(M{categoryColumnNo + 3}- M{categoryColumnNo - 2})*100", $"=(N{categoryColumnNo + 3}- N{categoryColumnNo - 2})*100", $"=(O{categoryColumnNo + 3}- 0{categoryColumnNo - 2})*100" };
        //            var additionalValues3 = new[] { yearData.First().Year.ToString(), "Proj Sales UL", $"=C{categoryColumnNo + 1}", $"=D{categoryColumnNo + 1}", $"=E{categoryColumnNo + 1}", $"=F{categoryColumnNo + 1}", $"=G{categoryColumnNo + 1}", $"=H{categoryColumnNo + 1}", $"=I{categoryColumnNo + 1}", $"=J{categoryColumnNo + 1}", $"=K{categoryColumnNo + 1}", $"=L{categoryColumnNo + 1}", $"=M{categoryColumnNo + 1}", $"=N{categoryColumnNo + 1}", $"=O{categoryColumnNo + 1}" };
        //            var additionalValues4 = new[] { yearData.First().Year.ToString(), "Proj $ % Chg", $"=C{categoryColumnNo + 5}/ C{categoryColumnNo - 3}-1", $"=D{categoryColumnNo + 5}/ D{categoryColumnNo - 3}-1", $"=E{categoryColumnNo + 5}/ E{categoryColumnNo - 3}-1", $"=F{categoryColumnNo + 5}/ F{categoryColumnNo - 3}-1", $"=G{categoryColumnNo + 5}/ G{categoryColumnNo - 3}-1", $"=H{categoryColumnNo + 5}/ H{categoryColumnNo - 3}-1", $"=I{categoryColumnNo + 5}/ I{categoryColumnNo - 3}-1", $"=J{categoryColumnNo + 5}/ J{categoryColumnNo - 3}-1", $"=K{categoryColumnNo + 5}/ K{categoryColumnNo - 3}-1", $"=L{categoryColumnNo + 5}/ L{categoryColumnNo - 3}-1", $"=M{categoryColumnNo + 5}/ M{categoryColumnNo - 3}-1", $"=N{categoryColumnNo + 5}/ N{categoryColumnNo - 3}-1", $"=O{categoryColumnNo + 5}/ {categoryColumnNo - 3}-1" };
        //            var additionalValues5 = new[] { yearData.First().Year.ToString(), "Proj UL $ Share", $"=C{categoryColumnNo + 5}/ C{categoryColumnNo}*100", $"=D{categoryColumnNo + 5}/ D{categoryColumnNo}*100", $"=E{categoryColumnNo + 5}/ E{categoryColumnNo}*100", $"=F{categoryColumnNo + 5}/ F{categoryColumnNo}*100", $"=G{categoryColumnNo + 5}/ G{categoryColumnNo}*100", $"=H{categoryColumnNo + 5}/ H{categoryColumnNo}*100", $"=I{categoryColumnNo + 5}/ I{categoryColumnNo}*100", $"=J{categoryColumnNo + 5}/ J{categoryColumnNo}*100", $"=K{categoryColumnNo + 5}/ K{categoryColumnNo}*100", $"=L{categoryColumnNo + 5}/ L{categoryColumnNo}*100", $"=M{categoryColumnNo + 5}/ M{categoryColumnNo}*100", $"=N{categoryColumnNo + 5}/ N{categoryColumnNo}*100", $"=O{categoryColumnNo + 5}/ O{categoryColumnNo}*100" };
        //            var additionalValues6 = new[] { yearData.First().Year.ToString(), "Proj BPS Chg", $"=(C{categoryColumnNo + 7} - C{categoryColumnNo - 2})*100", $"=(D{categoryColumnNo + 7} - D{categoryColumnNo - 2})*100", $"=(E{categoryColumnNo + 7} - E{categoryColumnNo - 2})*100", $"=(F{categoryColumnNo + 7} - F{categoryColumnNo - 2})*100", $"=(G{categoryColumnNo + 7} - G{categoryColumnNo - 2})*100", $"=(H{categoryColumnNo + 7} - H{categoryColumnNo - 2})*100", $"=(I{categoryColumnNo + 7} - I{categoryColumnNo - 2})*100", $"=(J{categoryColumnNo + 7} - J{categoryColumnNo - 2})*100", $"=(K{categoryColumnNo + 7} - K{categoryColumnNo - 2})*100", $"=(L{categoryColumnNo + 7} - L{categoryColumnNo - 2})*100", $"=(M{categoryColumnNo + 7} - M{categoryColumnNo - 2})*100", $"=(N{categoryColumnNo + 7} - N{categoryColumnNo - 2})*100", $"=(O{categoryColumnNo + 7} - O{categoryColumnNo - 2})*100" };
        //            var additionalValues7 = new[] { yearData.First().Year.ToString(), "Proj Bps vs Forecast", "0.0", "0.0", $"=(E{categoryColumnNo + 8} - E{categoryColumnNo + 4})", $"=(F{categoryColumnNo + 8} - F{categoryColumnNo + 4})", $"=(G{categoryColumnNo + 8} - G{categoryColumnNo + 4})", $"=(H{categoryColumnNo + 8} - H{categoryColumnNo + 4})", $"=(I{categoryColumnNo + 8} - I{categoryColumnNo + 4})", $"=(J{categoryColumnNo + 8} - J{categoryColumnNo + 4})", $"=(K{categoryColumnNo + 8} - K{categoryColumnNo + 4})", $"=(L{categoryColumnNo + 8} - L{categoryColumnNo + 4})", $"=(M{categoryColumnNo + 8} - M{categoryColumnNo + 4})", $"=(N{categoryColumnNo + 8} - N{categoryColumnNo + 4})", $"=(O{categoryColumnNo + 8} - O{categoryColumnNo + 4})" };
        //            summaryDataArray.Add(additionalValues);
        //            summaryDataArray.Add(additionalValues1);
        //            summaryDataArray.Add(additionalValues2);
        //            summaryDataArray.Add(additionalValues3);
        //            summaryDataArray.Add(additionalValues4);
        //            summaryDataArray.Add(additionalValues5);
        //            summaryDataArray.Add(additionalValues6);
        //            summaryDataArray.Add(additionalValues7);
        //            Dictionary<string, int> mergeList1 = new Dictionary<string, int>
        //            {
        //                { "row", categoryColumnNo -1 },
        //                { "col", 0 },
        //                { "rowspan",10 },
        //                { "colspan", 1 }
        //            };

        //            mergeArray.Add(mergeList1);

        //        }
        //        else
        //        {
        //            var additionalValues = new[] { yearData.First().Year.ToString(), "UL $ Share", $"=C{categoryColumnNo + 1}/ C{categoryColumnNo}*100", $"=D{categoryColumnNo + 1}/ D{categoryColumnNo}*100", $"=E{categoryColumnNo + 1}/ E{categoryColumnNo}*100", $"=F{categoryColumnNo + 1}/ F{categoryColumnNo}*100", $"=G{categoryColumnNo + 1}/ G{categoryColumnNo}*100", $"=H{categoryColumnNo + 1}/ H{categoryColumnNo}*100", $"=I{categoryColumnNo + 1}/ I{categoryColumnNo}*100", $"=J{categoryColumnNo + 1}/ J{categoryColumnNo}*100", $"=K{categoryColumnNo + 1}/ K{categoryColumnNo}*100", $"=L{categoryColumnNo + 1}/ L{categoryColumnNo}*100", $"=M{categoryColumnNo + 1}/ M{categoryColumnNo}*100", $"=N{categoryColumnNo + 1}/ N{categoryColumnNo}*100", $"=O{categoryColumnNo + 1}/ O{categoryColumnNo}*100" };
        //            summaryDataArray.Add(additionalValues);

        //        }

        //        //}
        //        var nullarray = new[] { "" };
        //        summaryDataArray.Add(nullarray);

        //        Dictionary<string, int> mergeList = new Dictionary<string, int>
        //        {
        //            { "row", categoryColumnNo - 1 },
        //            { "col", 0 },
        //            { "rowspan", 3 },
        //            { "colspan", 1 }
        //        };

        //        mergeArray.Add(mergeList);
        //        // clsCommonMethods.bindingQuatersData(dataByYear, categoryColumnNo);
        //        categoryColumnNo += 4;

        //    }

        //    var output = new SummaryDataDTO
        //    {
        //        Data = summaryDataArray,
        //        MergeData = mergeArray
        //    };

        //    return Ok(output);
        //}


        #endregion

        #region SummaryDataByFilters 
        //[HttpPost("SummaryDataByFilters")]
        //public async Task<ActionResult<IEnumerable<SummaryDataDTO>>> getSummaryDataByBrand([FromBody] SummaryDataByBrandDTO summaryData)
        //{
        //    try
        //    {
        //        //= IFERROR(E23 / E22 * 100, 0)a
        //        var getDataByBrand = await _context.GetProcedures().USP_GetSummaryDataByBrandAsync(string.Join(",", summaryData.subCategory),
        //                                                                           string.Join(",", summaryData.DataSource),
        //                                                                           string.Join(",", summaryData.Brands),
        //                                                                           string.Join(",", summaryData.Years));
        //        var summaryDataArray = new List<string[]>();
        //        List<Dictionary<string, int>> mergeArray = new List<Dictionary<string, int>>();
        //        var dataByBrand = getDataByBrand
        //                                    .GroupBy(d => d.Brand)
        //                                    .OrderBy(g => g.Key)
        //                                    .SelectMany(g => g
        //                                        .OrderBy(d => d.Year)
        //                                        .ThenBy(d => d.SubCategory)
        //                                        .ThenBy(d => d.SalesType));

        //        var categoryColumnNo = 2;
        //        var uniqueBrands = dataByBrand.Select(g => g.Brand).Distinct();
        //        foreach (var brand in uniqueBrands)
        //        {
        //            List<USP_GetSummaryDataByBrandResult> brandFilteredList = dataByBrand.Where(p => p.Brand == brand).ToList();
        //            var subCategoryDataBySalesType = brandFilteredList.GroupBy(d => new { d.SubCategory });

        //            foreach (var subData in subCategoryDataBySalesType)
        //            {
        //                var dataByYear = subData.GroupBy(d => d.Year).OrderBy(g => g.Key);

        //                foreach (var yearData in dataByYear)
        //                {
        //                    var yearDataBySalesType = yearData.GroupBy(d => new { d.Brand, d.SubCategory, d.Year });

        //                    foreach (var dataBySale in yearDataBySalesType)
        //                    {
        //                        if (summaryData.rollup.Length == 0)
        //                            return BadRequest("select rollup");
        //                        var headers = clsCommonMethods.headers(summaryData.rollup).Replace("Brand", brand).Split(',');
        //                        var result = new string[headers.Length + 2];
        //                        //result[0] = dataBySale.Key.SubCategory;
        //                        result[0] = string.Join(",", summaryData.DataSource);
        //                        result[1] = dataBySale.Key.SubCategory;

        //                        for (var i = 1; i < result.Length - 1; i++)
        //                        {
        //                            result[i + 1] = headers[i - 1];
        //                        }

        //                        summaryDataArray.Add(result);
        //                        int loopCount = 0;
        //                        foreach (var data in dataBySale)
        //                        {
        //                            var nestedValues = new[] { "", data.Year.ToString(), data.SalesType, data._1.ToString(), data._2.ToString(), data._3.ToString(), data._4.ToString(), data._5.ToString(), data._6.ToString(), data._7.ToString(), data._8.ToString(), data._9.ToString(), data._10.ToString(), data._11.ToString(), data._12.ToString(), data._13.ToString() };

        //                            if (dataByYear.First().Key != yearData.Key)
        //                            {
        //                                nestedValues = nestedValues.Concat(new[] { "" }).ToArray();
        //                                if (summaryData.rollup.Contains("Full Year"))
        //                                {
        //                                    var yearMatValues = new[] { $"=SUM($D{categoryColumnNo + loopCount}:$O{categoryColumnNo + loopCount},$P{categoryColumnNo - 5 + loopCount})" };
        //                                    nestedValues = nestedValues.Concat(yearMatValues).ToArray();
        //                                }
        //                                if (summaryData.rollup.Contains("Half MAT"))
        //                                {//= SUM($P8, D12: I12
        //                                    var yearMatValues = new[] { $"=SUM($D{categoryColumnNo + loopCount}:$I{categoryColumnNo + loopCount},$P{categoryColumnNo - 5 + loopCount})" ,
        //                                                                $"=SUM($J{categoryColumnNo + loopCount}:$O{categoryColumnNo + loopCount})"
        //                                                              };
        //                                    nestedValues = nestedValues.Concat(yearMatValues).ToArray();
        //                                }
        //                                if (summaryData.rollup.Contains("Quarter MAT"))
        //                                {
        //                                    var yearMatValues = new[] { $"=SUM($D{categoryColumnNo + loopCount}:$F{categoryColumnNo + loopCount},$P{categoryColumnNo - 5 + loopCount})" ,
        //                                                                $"=SUM($G{categoryColumnNo + loopCount}:$I{categoryColumnNo + loopCount})",$"=SUM($J{categoryColumnNo + loopCount}:$L{categoryColumnNo + loopCount})",
        //                                                                $"=SUM($M{categoryColumnNo + loopCount}:$O{categoryColumnNo + loopCount})"
        //                                                              };
        //                                    nestedValues = nestedValues.Concat(yearMatValues).ToArray();
        //                                }

        //                            }
        //                            summaryDataArray.Add(nestedValues);
        //                            loopCount++;  // Increment the loop count variable
        //                        }


        //                        if (dataByYear.Last().Key == yearData.Key)
        //                        {
        //                            string[][] fiscalMonthFormulas = new[]
        //                            {
        //                                new[] {"", yearData.First().Year.ToString() ?? string.Empty,
        //                                    "$ % Chg", $"=IFERROR(D{categoryColumnNo + 1}/ D{categoryColumnNo - 3}-1, 0)", $"=IFERROR(E{categoryColumnNo + 1}/ E{categoryColumnNo - 3}-1, 0)",
        //                                     $"=IFERROR(F{categoryColumnNo + 1}/ F{categoryColumnNo - 3}-1, 0)", $"=IFERROR(G{categoryColumnNo + 1}/ G{categoryColumnNo - 3}-1, 0)",
        //                                     $"=IFERROR(H{categoryColumnNo + 1}/ H{categoryColumnNo - 3}-1, 0)", $"=IFERROR(I{categoryColumnNo + 1}/ I{categoryColumnNo - 3}-1, 0)",
        //                                     $"=IFERROR(J{categoryColumnNo + 1}/ J{categoryColumnNo - 3}-1, 0)", $"=IFERROR(K{categoryColumnNo + 1}/ K{categoryColumnNo - 3}-1, 0)",
        //                                     $"=IFERROR(L{categoryColumnNo + 1}/ L{categoryColumnNo - 3}-1, 0)", $"=IFERROR(M{categoryColumnNo + 1}/ M{categoryColumnNo - 3}-1, 0)",
        //                                     $"=IFERROR(N{categoryColumnNo + 1}/ N{categoryColumnNo - 3}-1, 0)", $"=IFERROR(O{categoryColumnNo + 1}/ O{categoryColumnNo - 3}-1, 0)",
        //                                     $"=IFERROR(P{categoryColumnNo + 1}/ P{categoryColumnNo - 3}-1, 0)"
        //                                      },
        //                                new[] {"", yearData.First().Year.ToString() ?? string.Empty, "UL $ Share",  $"=IFERROR(D{categoryColumnNo + 1}/ D{categoryColumnNo}*100, 0)"
        //                                     , $"=IFERROR(E{categoryColumnNo + 1}/ E{categoryColumnNo}*100, 0)"
        //                                     , $"=IFERROR(F{categoryColumnNo + 1}/ F{categoryColumnNo}*100, 0)", $"=IFERROR(G{categoryColumnNo + 1}/ G{categoryColumnNo}*100, 0)"
        //                                     , $"=IFERROR(H{categoryColumnNo + 1}/ H{categoryColumnNo}*100, 0)", $"=IFERROR(I{categoryColumnNo + 1}/ I{categoryColumnNo}*100, 0)"
        //                                     , $"=IFERROR(J{categoryColumnNo + 1}/ J{categoryColumnNo}*100, 0)", $"=IFERROR(K{categoryColumnNo + 1}/ K{categoryColumnNo}*100, 0)"
        //                                     , $"=IFERROR(L{categoryColumnNo + 1}/ L{categoryColumnNo}*100, 0)", $"=IFERROR(M{categoryColumnNo + 1}/ M{categoryColumnNo}*100, 0)"
        //                                     , $"=IFERROR(N{categoryColumnNo + 1}/ N{categoryColumnNo}*100, 0)", $"=IFERROR(O{categoryColumnNo + 1}/ O{categoryColumnNo}*100, 0)"
        //                                     , $"=IFERROR(P{categoryColumnNo + 1}/ P{categoryColumnNo}*100, 0)"
        //                                      },
        //                                new[] {"", yearData.First().Year.ToString() ?? string.Empty, "UL BPS Chg", $"=IFERROR((D{categoryColumnNo + 3}- D{categoryColumnNo - 2})*100, 0)"
        //                                      , $"=IFERROR((E{categoryColumnNo + 3}- E{categoryColumnNo - 2})*100, 0)"
        //                                     , $"=IFERROR((F{categoryColumnNo + 3}- F{categoryColumnNo - 2})*100, 0)", $"=IFERROR((G{categoryColumnNo + 3}- G{categoryColumnNo - 2})*100, 0)"
        //                                     , $"=IFERROR((H{categoryColumnNo + 3}- H{categoryColumnNo - 2})*100, 0)", $"=IFERROR((I{categoryColumnNo + 3}- I{categoryColumnNo - 2})*100, 0)"
        //                                     , $"=IFERROR((J{categoryColumnNo + 3}- J{categoryColumnNo - 2})*100, 0)", $"=IFERROR((K{categoryColumnNo + 3}- K{categoryColumnNo - 2})*100, 0)"
        //                                     , $"=IFERROR((L{categoryColumnNo + 3}- L{categoryColumnNo - 2})*100, 0)", $"=IFERROR((M{categoryColumnNo + 3}- M{categoryColumnNo - 2})*100, 0)"
        //                                     , $"=IFERROR((N{categoryColumnNo + 3}- N{categoryColumnNo - 2})*100, 0)", $"=IFERROR((O{categoryColumnNo + 3}- O{categoryColumnNo - 2})*100, 0)"
        //                                     , $"=IFERROR((P{categoryColumnNo + 3}- P{categoryColumnNo - 2})*100, 0)"
        //                                      },
        //                                new[] {"", yearData.First().Year.ToString() ?? string.Empty, "Proj Sales UL", $"=IFERROR(D{categoryColumnNo + 1}, 0)"
        //                                     , $"=IFERROR(E{categoryColumnNo + 1}, 0)", $"=IFERROR(F{categoryColumnNo + 1}, 0)"
        //                                     , $"=IFERROR(G{categoryColumnNo + 1}, 0)", $"=IFERROR(H{categoryColumnNo + 1}, 0)", $"=IFERROR(I{categoryColumnNo + 1}, 0)"
        //                                     , $"=IFERROR(J{categoryColumnNo + 1}, 0)", $"=IFERROR(K{categoryColumnNo + 1}, 0)", $"=IFERROR(L{categoryColumnNo + 1}, 0)"
        //                                     , $"=IFERROR(M{categoryColumnNo + 1}, 0)", $"=IFERROR(N{categoryColumnNo + 1}, 0)", $"=IFERROR(O{categoryColumnNo + 1}, 0)"
        //                                     , $"=IFERROR(P{categoryColumnNo + 1}, 0)"
        //                                      },
        //                                new[] {"", yearData.First().Year.ToString() ?? string.Empty, "Proj $ % Chg", $"=IFERROR(D{categoryColumnNo + 5}/ D{categoryColumnNo - 3}-1, 0)"
        //                                      , $"=IFERROR(E{categoryColumnNo + 5}/ E{categoryColumnNo - 3}-1, 0)"
        //                                      , $"=IFERROR(F{categoryColumnNo + 5}/ F{categoryColumnNo - 3}-1, 0)", $"=IFERROR(G{categoryColumnNo + 5}/ G{categoryColumnNo - 3}-1, 0)"
        //                                      , $"=IFERROR(H{categoryColumnNo + 5}/ H{categoryColumnNo - 3}-1, 0)", $"=IFERROR(I{categoryColumnNo + 5}/ I{categoryColumnNo - 3}-1, 0)"
        //                                      , $"=IFERROR(J{categoryColumnNo + 5}/ J{categoryColumnNo - 3}-1, 0)", $"=IFERROR(K{categoryColumnNo + 5}/ K{categoryColumnNo - 3}-1, 0)"
        //                                      , $"=IFERROR(L{categoryColumnNo + 5}/ L{categoryColumnNo - 3}-1, 0)", $"=IFERROR(M{categoryColumnNo + 5}/ M{categoryColumnNo - 3}-1, 0)"
        //                                      , $"=IFERROR(N{categoryColumnNo + 5}/ N{categoryColumnNo - 3}-1, 0)", $"=IFERROR(O{categoryColumnNo + 5}/ {categoryColumnNo - 3}-1, 0)"
        //                                      , $"=IFERROR(P{categoryColumnNo + 5}/ P{categoryColumnNo - 3}-1, 0)"
        //                                      },
        //                                new[] {"", yearData.First().Year.ToString() ?? string.Empty, "Proj UL $ Share", $"=IFERROR(D{categoryColumnNo + 5}/ D{categoryColumnNo}*100, 0)"
        //                                      , $"=IFERROR(E{categoryColumnNo + 5}/ E{categoryColumnNo}*100, 0)"
        //                                      , $"=IFERROR(F{categoryColumnNo + 5}/ F{categoryColumnNo}*100, 0)", $"=IFERROR(G{categoryColumnNo + 5}/ G{categoryColumnNo}*100, 0)"
        //                                      , $"=IFERROR(H{categoryColumnNo + 5}/ H{categoryColumnNo}*100, 0)", $"=IFERROR(I{categoryColumnNo + 5}/ I{categoryColumnNo}*100, 0)"
        //                                      , $"=IFERROR(J{categoryColumnNo + 5}/ J{categoryColumnNo}*100, 0)", $"=IFERROR(K{categoryColumnNo + 5}/ K{categoryColumnNo}*100, 0)"
        //                                      , $"=IFERROR(L{categoryColumnNo + 5}/ L{categoryColumnNo}*100, 0)", $"=IFERROR(M{categoryColumnNo + 5}/ M{categoryColumnNo}*100, 0)"
        //                                      , $"=IFERROR(N{categoryColumnNo + 5}/ N{categoryColumnNo}*100, 0)", $"=IFERROR(O{categoryColumnNo + 5}/ O{categoryColumnNo}*100, 0)"
        //                                      , $"=IFERROR(P{categoryColumnNo + 5}/ P{categoryColumnNo}*100, 0)"
        //                                      },
        //                                new[] {"", yearData.First().Year.ToString() ?? string.Empty, "Proj BPS Chg", $"=(D{categoryColumnNo + 7} - D{categoryColumnNo - 2})*100"
        //                                      , $"=(E{categoryColumnNo + 7} - E{categoryColumnNo - 2})*100"
        //                                      , $"=(F{categoryColumnNo + 7} - F{categoryColumnNo - 2})*100", $"=(G{categoryColumnNo + 7} - G{categoryColumnNo - 2})*100"
        //                                      , $"=(H{categoryColumnNo + 7} - H{categoryColumnNo - 2})*100", $"=(I{categoryColumnNo + 7} - I{categoryColumnNo - 2})*100"
        //                                      , $"=(J{categoryColumnNo + 7} - J{categoryColumnNo - 2})*100", $"=(K{categoryColumnNo + 7} - K{categoryColumnNo - 2})*100"
        //                                      , $"=(L{categoryColumnNo + 7} - L{categoryColumnNo - 2})*100", $"=(M{categoryColumnNo + 7} - M{categoryColumnNo - 2})*100"
        //                                      , $"=(N{categoryColumnNo + 7} - N{categoryColumnNo - 2})*100", $"=(O{categoryColumnNo + 7} - O{categoryColumnNo - 2})*100"
        //                                      , $"=(P{categoryColumnNo + 7} - P{categoryColumnNo - 2})*100"
        //                                      },
        //                                new[] {"", yearData.First().Year.ToString() ?? string.Empty, "Proj Bps vs Forecast", "0.0", "0.0"
        //                                        , $"=IFERROR(F{categoryColumnNo + 8} - F{categoryColumnNo + 4}, 0)", $"=IFERROR(G{categoryColumnNo + 8} - G{categoryColumnNo + 4}, 0)"
        //                                        , $"=IFERROR(H{categoryColumnNo + 8} - H{categoryColumnNo + 4}, 0)", $"=IFERROR(I{categoryColumnNo + 8} - I{categoryColumnNo + 4}, 0)"
        //                                        , $"=IFERROR(J{categoryColumnNo + 8} - J{categoryColumnNo + 4}, 0)", $"=IFERROR(K{categoryColumnNo + 8} - K{categoryColumnNo + 4}, 0)"
        //                                        , $"=IFERROR(L{categoryColumnNo + 8} - L{categoryColumnNo + 4}, 0)", $"=IFERROR(M{categoryColumnNo + 8} - M{categoryColumnNo + 4}, 0)"
        //                                        , $"=IFERROR(N{categoryColumnNo + 8} - N{categoryColumnNo + 4}, 0)", $"=IFERROR(O{categoryColumnNo + 8} - O{categoryColumnNo + 4}, 0)"
        //                                        , $"=IFERROR(P{categoryColumnNo + 8} - P{categoryColumnNo + 4}, 0)" }
        //                            };

        //                            if (summaryData.rollup.Count() != 0)
        //                            {
        //                                var NullArray = new[] {
        //                                    new[] {"" },
        //                                    new[] {"" },
        //                                    new[] {"" },
        //                                    new[] {"" },
        //                                    new[] {"" },
        //                                    new[] {"" },
        //                                    new[] {"" },
        //                                    new[] {"" },
        //                                };
        //                                for (int i = 0; i < fiscalMonthFormulas.Length; i++)
        //                                {
        //                                    fiscalMonthFormulas[i] = fiscalMonthFormulas[i].Concat(NullArray[i]).ToArray();
        //                                }
        //                            }



        //                            if (summaryData.rollup.Contains("Full Year"))
        //                            {
        //                                var yearMatValues = new[]
        //                                {
        //                                  new [] { $"=IFERROR(R{categoryColumnNo + 1}/ R{categoryColumnNo - 3}-1, 0)" },
        //                                  new [] { $"=IFERROR(R{categoryColumnNo + 1}/ R{categoryColumnNo}*100, 0)" },
        //                                  new [] { $"=IFERROR((R{categoryColumnNo + 3}- R{categoryColumnNo - 2})*100, 0)" },
        //                                  new [] { $"=IFERROR(R{categoryColumnNo + 1}, 0)" },
        //                                  new [] { $"=IFERROR(R{categoryColumnNo + 5}/ R{categoryColumnNo - 3}-1, 0)" },
        //                                  new [] { $"=IFERROR(R{categoryColumnNo + 5}/ R{categoryColumnNo}*100, 0)" },
        //                                  new [] { $"=(R{categoryColumnNo + 7} - R{categoryColumnNo - 2})*100" },
        //                                  new [] { $"=IFERROR(R{categoryColumnNo + 8} - R{categoryColumnNo + 4}, 0)"}
        //                                };
        //                                for (int i = 0; i < fiscalMonthFormulas.Length; i++)
        //                                {
        //                                    fiscalMonthFormulas[i] = fiscalMonthFormulas[i].Concat(yearMatValues[i]).ToArray();
        //                                }
        //                            }
        //                            if (summaryData.rollup.Contains("Half MAT"))
        //                            {
        //                                var yearMatValues = new[]
        //                                {
        //                                  new [] { $"=IFERROR(S{categoryColumnNo + 1}/ S{categoryColumnNo - 3}-1, 0)" ,$"=IFERROR(T{categoryColumnNo + 1}/ T{categoryColumnNo - 3}-1, 0)" },
        //                                  new [] { $"=IFERROR(S{categoryColumnNo + 1}/ S{categoryColumnNo}*100, 0)" , $"=IFERROR(T{categoryColumnNo + 1}/ T{categoryColumnNo}*100, 0)" },
        //                                  new [] { $"=IFERROR((S{categoryColumnNo + 3}- S{categoryColumnNo - 2})*100, 0)",$"=IFERROR((T{categoryColumnNo + 3}- T{categoryColumnNo - 2})*100, 0)" },
        //                                  new [] { $"=IFERROR(S{categoryColumnNo + 1}, 0)" ,$"=IFERROR(T{categoryColumnNo + 1}, 0)" },
        //                                  new [] { $"=IFERROR(S{categoryColumnNo + 5}/ S{categoryColumnNo - 3}-1, 0)",$"=IFERROR(T{categoryColumnNo + 5}/ T{categoryColumnNo - 3}-1, 0)" },
        //                                  new [] { $"=IFERROR(S{categoryColumnNo + 5}/ S{categoryColumnNo}*100, 0)",$"=IFERROR(T{categoryColumnNo + 5}/ T{categoryColumnNo}*100, 0)" },
        //                                  new [] { $"=(S{categoryColumnNo + 7} - S{categoryColumnNo - 2})*100",$"=(S{categoryColumnNo + 7} - T{categoryColumnNo - 2})*100" },
        //                                  new [] { $"=IFERROR(S{categoryColumnNo + 8} - S{categoryColumnNo + 4}, 0)", $"=IFERROR(T{categoryColumnNo + 8} - T{categoryColumnNo + 4}, 0)"}
        //                                };
        //                                for (int i = 0; i < fiscalMonthFormulas.Length; i++)
        //                                {
        //                                    fiscalMonthFormulas[i] = fiscalMonthFormulas[i].Concat(yearMatValues[i]).ToArray();
        //                                }
        //                            }
        //                            if (summaryData.rollup.Contains("Quarter MAT"))
        //                            {
        //                                var yearMatValues = new[]
        //                                {
        //                                  new [] { $"=IFERROR(u{categoryColumnNo + 1}/ u{categoryColumnNo - 3}-1, 0)" ,$"=IFERROR(V{categoryColumnNo + 1}/ V{categoryColumnNo - 3}-1, 0)",
        //                                           $"=IFERROR(W{categoryColumnNo + 1}/ W{categoryColumnNo - 3}-1, 0)" ,$"=IFERROR(X{categoryColumnNo + 1}/ X{categoryColumnNo - 3}-1, 0)"},
        //                                  new [] { $"=IFERROR(u{categoryColumnNo + 1}/ u{categoryColumnNo}*100, 0)" , $"=IFERROR(V{categoryColumnNo + 1}/ V{categoryColumnNo}*100, 0)",
        //                                           $"=IFERROR(W{categoryColumnNo + 1}/ W{categoryColumnNo}*100, 0)" , $"=IFERROR(X{categoryColumnNo + 1}/ X{categoryColumnNo}*100, 0)"},
        //                                  new [] { $"=IFERROR((u{categoryColumnNo + 3}- u{categoryColumnNo - 2})*100, 0)",$"=IFERROR((V{categoryColumnNo + 3}- V{categoryColumnNo - 2})*100, 0)",
        //                                           $"=IFERROR((W{categoryColumnNo + 3}- W{categoryColumnNo - 2})*100, 0)",$"=IFERROR((X{categoryColumnNo + 3}- X{categoryColumnNo - 2})*100, 0)"},
        //                                  new [] { $"=IFERROR(u{categoryColumnNo + 1}, 0)" ,$"=IFERROR(v{categoryColumnNo + 1}, 0)",$"=IFERROR(W{categoryColumnNo + 1}, 0)" ,$"=IFERROR(W{categoryColumnNo + 1}, 0)" },
        //                                  new [] { $"=IFERROR(u{categoryColumnNo + 5}/ u{categoryColumnNo - 3}-1, 0)",$"=IFERROR(V{categoryColumnNo + 5}/ V{categoryColumnNo - 3}-1, 0)" ,
        //                                           $"=IFERROR(W{categoryColumnNo + 5}/ W{categoryColumnNo - 3}-1, 0)",$"=IFERROR(X{categoryColumnNo + 5}/ X{categoryColumnNo - 3}-1, 0)" },
        //                                  new [] { $"=IFERROR(u{categoryColumnNo + 5}/ u{categoryColumnNo}*100, 0)",$"=IFERROR(V{categoryColumnNo + 5}/ V{categoryColumnNo}*100, 0)" ,
        //                                           $"=IFERROR(W{categoryColumnNo + 5}/ W{categoryColumnNo}*100, 0)",$"=IFERROR(X{categoryColumnNo + 5}/ X{categoryColumnNo}*100, 0)" },
        //                                  new [] { $"=(u{categoryColumnNo + 7} - u{categoryColumnNo - 2})*100",$"=(V{categoryColumnNo + 7} - V{categoryColumnNo - 2})*100" ,
        //                                           $"=(W{categoryColumnNo + 7} - W{categoryColumnNo - 2})*100",$"=(X{categoryColumnNo + 7} - X{categoryColumnNo - 2})*100" },
        //                                  new [] { $"=IFERROR(u{categoryColumnNo + 8} - u{categoryColumnNo + 4}, 0)", $"=IFERROR(V{categoryColumnNo + 8} - V{categoryColumnNo + 4}, 0)",
        //                                           $"=IFERROR(W{categoryColumnNo + 8} - W{categoryColumnNo + 4}, 0)", $"=IFERROR(X{categoryColumnNo + 8} - W{categoryColumnNo + 4}, 0)"}
        //                                };
        //                                for (int i = 0; i < fiscalMonthFormulas.Length; i++)
        //                                {
        //                                    fiscalMonthFormulas[i] = fiscalMonthFormulas[i].Concat(yearMatValues[i]).ToArray();
        //                                }
        //                            }


        //                            foreach (var Values in fiscalMonthFormulas)
        //                            {
        //                                summaryDataArray.Add(Values);
        //                            }
        //                        }
        //                        else
        //                        {
        //                            var additionalValues = new[] { "", yearData.First().Year.ToString() ?? string.Empty, "UL $ Share", $"=P{categoryColumnNo + 1}/ P{categoryColumnNo}*100"
        //                                            , $"=D{categoryColumnNo + 1}/ D{categoryColumnNo}*100", $"=E{categoryColumnNo + 1}/ E{categoryColumnNo}*100"
        //                                            , $"=F{categoryColumnNo + 1}/ F{categoryColumnNo}*100", $"=G{categoryColumnNo + 1}/ G{categoryColumnNo}*100"
        //                                            , $"=H{categoryColumnNo + 1}/ H{categoryColumnNo}*100", $"=I{categoryColumnNo + 1}/ I{categoryColumnNo}*100"
        //                                            , $"=J{categoryColumnNo + 1}/ J{categoryColumnNo}*100", $"=K{categoryColumnNo + 1}/ K{categoryColumnNo}*100"
        //                                            , $"=L{categoryColumnNo + 1}/ L{categoryColumnNo}*100", $"=M{categoryColumnNo + 1}/ M{categoryColumnNo}*100"
        //                                            , $"=N{categoryColumnNo + 1}/ N{categoryColumnNo}*100", $"=O{categoryColumnNo + 1}/ O{categoryColumnNo}*100" };


        //                            if (dataByYear.First().Key != yearData.Key)
        //                            {
        //                                additionalValues = additionalValues.Concat(new[] { "" }).ToArray();
        //                                if (summaryData.rollup.Contains("Full Year"))
        //                                {
        //                                    var yearMatValues = new[] { $"=R{categoryColumnNo + 1}/ R{categoryColumnNo}*100" };
        //                                    additionalValues = additionalValues.Concat(yearMatValues).ToArray();
        //                                }
        //                                if (summaryData.rollup.Contains("Half MAT"))
        //                                {
        //                                    var yearMatValues = new[] { $"=S{categoryColumnNo + 1}/ S{categoryColumnNo}*100", $"=T{categoryColumnNo + 1}/ T{categoryColumnNo}*100" };
        //                                    additionalValues = additionalValues.Concat(yearMatValues).ToArray();
        //                                }
        //                                if (summaryData.rollup.Contains("Quarter MAT"))
        //                                {
        //                                    var yearMatValues = new[] {$"=U{categoryColumnNo + 1}/ U{categoryColumnNo}*100", $"=V{categoryColumnNo + 1}/ V{categoryColumnNo}*100" ,
        //                                                               $"=W{categoryColumnNo + 1}/ W{categoryColumnNo}*100", $"=X{categoryColumnNo + 1}/ X{categoryColumnNo}*100"};
        //                                    additionalValues = additionalValues.Concat(yearMatValues).ToArray();
        //                                }

        //                            }
        //                            summaryDataArray.Add(additionalValues);
        //                        }
        //                        var BrandmergeList = new Dictionary<string, int>
        //                            {
        //                                { "row", categoryColumnNo - 1 },
        //                                { "col", 1},
        //                                { "rowspan", (dataByYear.Last().Key == yearData.Key) ? 10 : 3 },
        //                                { "colspan", 1 }
        //                            };
        //                        var DataSourcemergeList = new Dictionary<string, int>
        //                            {
        //                                { "row", categoryColumnNo - 1 },
        //                                { "col", 0},
        //                                { "rowspan", (dataByYear.Last().Key == yearData.Key) ? 10 : 3 },
        //                                { "colspan", 1 }
        //                            };
        //                        mergeArray.Add(BrandmergeList);
        //                        mergeArray.Add(DataSourcemergeList);
        //                        categoryColumnNo += (dataByYear.Last().Key == yearData.Key) ? 12 : 5;
        //                        summaryDataArray.Add(new[] { "" });
        //                    }
        //                }
        //            }
        //        }
        //        var output = new SummaryDataDTO
        //        {
        //            Data = summaryDataArray,
        //            MergeData = mergeArray
        //        };

        //        return Ok(output);
        //    }
        //    catch (Exception ex)
        //    {
        //        return BadRequest(ex.Message);
        //    }

        //}

        #endregion

        #region extractFileUpload 

        [HttpPost("extractFileUpload")]
        public IActionResult UploadFile(IFormFile file)
        {
            // Check if a file was provided
            if (file == null || file.Length == 0)
                return BadRequest("No file provided");

            //// Check if the file is an Excel file
            //if (!file.FileName.EndsWith(".xlsx"))
            //    return BadRequest("Invalid file format. Only .xlsx files are allowed.");

            try
            {
                // Read the Excel file using ClosedXML
                using (var workbook = new XLWorkbook(file.OpenReadStream()))
                {
                    var worksheet = workbook.Worksheet(1);
                    var range = worksheet.RangeUsed();

                    var rowCount = range.RowCount();
                    var columnCount = range.ColumnCount();

                    var result = new List<Dictionary<string, string>>();

                    // Convert each row to a dictionary
                    for (int row = 2; row <= rowCount; row++)
                    {
                        var dict = new Dictionary<string, string>();
                        for (int col = 1; col <= columnCount; col++)
                        {
                            var key = range.Cell(1, col).Value.ToString();
                            var value = range.Cell(row, col).Value.ToString();
                            dict[key] = value;
                        }
                        result.Add(dict);
                    }

                    // Convert the result to JSON
                    var json = JsonConvert.SerializeObject(result);

                    return Ok(json);
                }
            }
            catch (Exception ex)
            {
                return BadRequest("Error occurred while processing the file: " + ex.Message);
            }
        }

        #endregion
    }
}
