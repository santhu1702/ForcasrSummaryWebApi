﻿// <auto-generated> This file has been auto generated by EF Core Power Tools. </auto-generated>
using ForcasrSummaryWebApi.MetaData;
using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Data;
using System.Threading;
using System.Threading.Tasks;

namespace ForcasrSummaryWebApi.MetaData
{
    public partial class budgetForecastContext
    {
        private IbudgetForecastContextProcedures _procedures;

        public virtual IbudgetForecastContextProcedures Procedures
        {
            get
            {
                if (_procedures is null) _procedures = new budgetForecastContextProcedures(this);
                return _procedures;
            }
            set
            {
                _procedures = value;
            }
        }

        public IbudgetForecastContextProcedures GetProcedures()
        {
            return Procedures;
        }

        protected void OnModelCreatingGeneratedProcedures(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<USP_GetSummaryDataResult>().HasNoKey().ToView(null);
            modelBuilder.Entity<USP_GetSummaryDataByBrandResult>().HasNoKey().ToView(null);
        }
    }

    public partial class budgetForecastContextProcedures : IbudgetForecastContextProcedures
    {
        private readonly budgetForecastContext _context;

        public budgetForecastContextProcedures(budgetForecastContext context)
        {
            _context = context;
        }

        public virtual async Task<List<USP_GetSummaryDataResult>> USP_GetSummaryDataAsync(OutputParameter<int> returnValue = null, CancellationToken cancellationToken = default)
        {
            var parameterreturnValue = new SqlParameter
            {
                ParameterName = "returnValue",
                Direction = System.Data.ParameterDirection.Output,
                SqlDbType = System.Data.SqlDbType.Int,
            };

            var sqlParameters = new []
            {
                parameterreturnValue,
            };
            var _ = await _context.SqlQueryAsync<USP_GetSummaryDataResult>("EXEC @returnValue = [dbo].[USP_GetSummaryData]", sqlParameters, cancellationToken);

            returnValue?.SetValue(parameterreturnValue.Value);

            return _;
        }

        public virtual async Task<List<USP_GetSummaryDataByBrandResult>> USP_GetSummaryDataByBrandAsync(string SubCatagoery, OutputParameter<int> returnValue = null, CancellationToken cancellationToken = default)
        {
            var parameterreturnValue = new SqlParameter
            {
                ParameterName = "returnValue",
                Direction = System.Data.ParameterDirection.Output,
                SqlDbType = System.Data.SqlDbType.Int,
            };

            var sqlParameters = new []
            {
                new SqlParameter
                {
                    ParameterName = "SubCatagoery",
                    Size = 2000,
                    Value = SubCatagoery ?? Convert.DBNull,
                    SqlDbType = System.Data.SqlDbType.NVarChar,
                },
                parameterreturnValue,
            };
            var _ = await _context.SqlQueryAsync<USP_GetSummaryDataByBrandResult>("EXEC @returnValue = [dbo].[USP_GetSummaryDataByBrand] @SubCatagoery", sqlParameters, cancellationToken);

            returnValue?.SetValue(parameterreturnValue.Value);

            return _;
        }
    }
}
