using Microsoft.AspNetCore.Mvc;
using ForcasrSummaryWebApi.MetaData;
using System.Data;
using Microsoft.EntityFrameworkCore;

namespace ForcasrSummaryWebApi.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class ExcelController : ControllerBase
    {
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
        public async Task<ActionResult<USP_GetSummaryDataResult>> getSummaryData()
        {
            return Ok(await _context.GetProcedures().USP_GetSummaryDataAsync());
        }
        [HttpGet("SummaryDataByBrand")]
        public async Task<ActionResult<USP_GetSummaryDataByBrandResult>> getSummaryDataByBrand(string SubCatagoery)
        {
            return Ok(await _context.GetProcedures().USP_GetSummaryDataByBrandAsync(SubCatagoery));
        }
    }
}