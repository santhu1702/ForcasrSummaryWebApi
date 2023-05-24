using System;
using System.Text;
using ForcasrSummaryWebApi.MetaData;

namespace ForcasrSummaryWebApi.CommonMethods
{
    public static class clsCommonMethods
    {
        public static string _headers = "Brand,C1,C2,C3,C4,C5,C6,C7,C8,C9,C10,C11,C12,C13";
        public static string _quarterMATHeaders = "Q1,Q2,Q3,Q4";
        public static string _halfYearMATHeaders = "H1,H2";
        public static string _fullYearMATHeaders = "FY MAT";

        public static string headers(string[] selectedMat)
        {
            var stringBuilder = new StringBuilder(_headers);
            stringBuilder.Append(",");

            if (selectedMat.Contains("Full Year"))
            {
                stringBuilder.Append(","); 
                stringBuilder.Append(_fullYearMATHeaders);
            }

            if (selectedMat.Contains("Half MAT"))
            {
                stringBuilder.Append(","); 
                stringBuilder.Append(_halfYearMATHeaders);
            }

            if (selectedMat.Contains("Quarter MAT"))
            {
                stringBuilder.Append(","); 
                stringBuilder.Append(_quarterMATHeaders);
            }

            return stringBuilder.ToString();
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

