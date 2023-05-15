using System;
using ForcasrSummaryWebApi.MetaData;

namespace ForcasrSummaryWebApi.CommonMethods
{
	public class clsCommonMethods
	{
        public static string _quaterHeaders = "FY MAT,H1,H2, Q1,Q2,Q3,Q4";


        public string bindingQuatersData(IOrderedEnumerable<IGrouping<long?, USP_GetSummaryDataResult>> year, int categoryColumnNo)
            {
                var quaterHeaders = _quaterHeaders.Split(',');
                var result = new string[quaterHeaders.Length + 1];
                result[0] = "Body Cleansing";

                for (var i = 1; i < result.Length; i++)
                {
                    result[i] = quaterHeaders[i - 1];
                }
                var quaterDataArray = new List<string[]> { result };
                var yearData = year;

                foreach (var years in year)
                {
                    if (years.First() == yearData)
                    {
                        var values = new[] { "", "", "", "", "", "", "" };
                        quaterDataArray.Add(values);
                    }
                    else
                    {
                    }

                }
                return "";
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
        }
	
}

