using System.Data;
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
        private readonly budgetForecastContext? _context;
        public ExcelController(budgetForecastContext? context)
        {
            _context = context;
        }



        [HttpGet("SubCatagoery")]
        public async Task<ActionResult<SubCategoryDTO>> GetSubCategoryAsync()
        {
            try
            {
                var categories = await _context.SummaryForcast
                    .Select(x => x.SubCategory)
                    .Distinct()
                    .ToListAsync();
                return Ok(categories.ToList());
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

            foreach (var yearData in dataByYear)
            {
                foreach (var data in yearData)
                {
                    var nestedvalues = new[] { data.Year.ToString(), data.SalesType, data._1.ToString(), data._2.ToString(), data._3.ToString(), data._4.ToString(), data._5.ToString(), data._6.ToString(), data._7.ToString(), data._8.ToString(), data._9.ToString(), data._10.ToString(), data._11.ToString(), data._12.ToString(), data._13.ToString() };
                    summaryDataArray.Add(nestedvalues);
                }

                var additionalValues = new[] { yearData.First().Year.ToString(), "UL $ Share", $"= SUM(C{categoryColumnNo + 1}: C{categoryColumnNo + 1})", $"= SUM(D{categoryColumnNo}: D{categoryColumnNo + 1})", $"= SUM(F{categoryColumnNo}: F{categoryColumnNo + 1})", $"= SUM(G{categoryColumnNo}: G{categoryColumnNo + 1})", $"= SUM(H{categoryColumnNo}: H{categoryColumnNo + 1})", $"= SUM(H{categoryColumnNo}: H{categoryColumnNo + 1})", $"= SUM(I{categoryColumnNo}: I{categoryColumnNo + 1})", $"= SUM(J{categoryColumnNo}: J{categoryColumnNo + 1})", $"= SUM(K{categoryColumnNo}: K{categoryColumnNo + 1})", $"= SUM(L{categoryColumnNo}: L{categoryColumnNo + 1})", $"= SUM(M{categoryColumnNo}: M{categoryColumnNo + 1})", $"= SUM(N{categoryColumnNo}: N{categoryColumnNo + 1})", $"= SUM(O{categoryColumnNo}: O{categoryColumnNo + 1})" };
                summaryDataArray.Add(additionalValues);
                if (dataByYear.Last().Key != yearData.Key)
                {
                    //need to implement
                }
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
                categoryColumnNo += 4;

            }

            var output = new
            {
                data = summaryDataArray,
                summaryDataArray = mergeArray
            };

            return Ok(output);
        }


        public static string getColumnName(int columnNumber)
        {
            const string letters = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
            string columnName = "";

            while (columnNumber > 0)
            {
                columnName = letters[(columnNumber - 1) % 26] + columnName;
                columnNumber = (columnNumber - 1) / 26;
            }

            return columnName;
        }
        [HttpGet("SummaryDataByBrand")]
        public async Task<ActionResult<USP_GetSummaryDataByBrandResult>> getSummaryDataByBrand(string SubCatagoery)
        {
            return Ok(await _context.GetProcedures().USP_GetSummaryDataByBrandAsync(SubCatagoery));
        }


    }
}