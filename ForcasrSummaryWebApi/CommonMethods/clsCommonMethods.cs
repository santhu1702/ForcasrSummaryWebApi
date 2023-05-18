using System;
using ForcasrSummaryWebApi.MetaData;

namespace ForcasrSummaryWebApi.CommonMethods
{
    public static class clsCommonMethods
    {
        public static string _quaterHeaders = "FY MAT,H1,H2, Q1,Q2,Q3,Q4";


        public static string bindingQuatersData(IOrderedEnumerable<IGrouping<long?, USP_GetSummaryDataResult>> year, int categoryColumnNo)
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

        public static List<string[]> getsubbrand(List<USP_GetSummaryDataByBrandResult> dataList, string headers, string quaterHeaders, List<string[]> summaryDataArray, string subCategoryName)
        {
            try
            {

            var dataByBrand = dataList.GroupBy(d => d.Brand)
                                        .OrderBy(g => g.Key == "Total" ? 0 : 1);
            var categoryColumnNo = 25;
            summaryDataArray.Add(new[] { "" });
            foreach (var brandData in dataByBrand)
            {
                var brand = brandData.Key;
                var summaryHeaders = headers.Replace("Category", Convert.ToString(subCategoryName)).Replace("Brand", brandData.Key).Split(',');
                var result = new string[summaryHeaders.Length + 1];
                result[0] = subCategoryName;
                    for (var i = 1; i < result.Length; i++)
                {
                    result[i] = summaryHeaders[i - 1];
                }
                summaryDataArray.Add(result);

                var dataByYear = brandData.GroupBy(d => d.Year).OrderBy(g => g.Key);
                List<Dictionary<string, int>> mergeArray = new List<Dictionary<string, int>>();

                foreach (var yearData in dataByYear)
                {
                    var year = yearData.Key;
                    foreach (var data in yearData)
                    {
                        var nestedvalues = new[] { data.Year.ToString(), data.SalesType, data._1.ToString(), data._2.ToString(), data._3.ToString(), data._4.ToString(), data._5.ToString(), data._6.ToString(), data._7.ToString(), data._8.ToString(), data._9.ToString(), data._10.ToString(), data._11.ToString(), data._12.ToString(), data._13.ToString() };
                        summaryDataArray.Add(nestedvalues);
                    }
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
                }


            }

            return summaryDataArray;

            }
            catch(Exception ex)
            {
                throw(ex);  
            }
        }

    }
}

