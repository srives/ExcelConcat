using System;
using System.Collections.Generic;
using System.Linq;

/// <summary>
///  Create an Excel SQL statement out of the header row and the row below it
///     Insert into table (a,b,c,d,...) VALUES (x, y, z...)
///  
/// </summary>
namespace XLSConcat
{
    class Program
    {

        /// <summary>
        /// Convert an integer to an EXCEL column name
        /// 
        /// 1 => A
        /// 2 => B
        /// ...
        /// 27 => AA
        /// 28 => AB
        /// 
        /// </summary>
        /// <param name="columnNumber"></param>
        /// <returns></returns>
        static public string GetExcelColName(int columnNumber)
        {
            int dividend = columnNumber;
            string columnName = String.Empty;
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnName = Convert.ToChar(65 + modulo).ToString() + columnName;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnName;
        }

        static void Main(string[] args)
        {
            // Assuming your spreadsheet starts with headers in A1...<NN>, and data below that
            // YOU will insert a fist column into your spreadsheet, and then paste the output of this program into A2
            //    After that, duplicate that cell in all the rows in column A

                        
            int columnCount = 31; // how many columns in the table to use
            int startingCol = 2; // 2 = B, this will make the script use B1 as the first column header title.
            string table = "CMC_NWPR_RELATION";
            string answer = $"=CONCAT(\"INSERT INTO {table} (\", ";
            bool columnNamesHaveSpaces = false; // if set to true, I will do a "" around each column header name
            bool debug = false;            


            // Every CHAR type column gets quoted and trimed
            List<int> stringColumns = new List<int> {2,3,4,5,8,9,13,14,15,16,17,18,21,22,24,26,27,28};

            // Date columns get a conversion function (you can change the format of DT values)
            List<int> dateColumns = new List<int> { 6, 7, 19, 25, 31 };
            string dateTimeFormat = "YYYY-MM-DD";

            for (int i = startingCol; i <= columnCount; i++)
            {
                //if (debug)
                //    answer += $"\"{i}\", \",\", ";
                if (columnNamesHaveSpaces)
                    answer += $"CHAR(34), ${GetExcelColName(i)}$1, char(34)";
                else
                    answer += $"${GetExcelColName(i)}$1";

                if (i != columnCount)
                    answer += ", \",\", ";
                else
                    answer += ", ";
            }

            answer += $"\") VALUES (\", ";

            for (int i = startingCol; i <= columnCount; i++)
            {
                if (debug)
                   answer += $"\"{i}\", \",\", ";
                if (dateColumns.Contains(i))
                    answer += $"CHAR(34), TEXT({GetExcelColName(i)}2, \"{dateTimeFormat}\"), char(34)";
                else if (stringColumns.Contains(i))
                    answer += $"CHAR(34), TRIM({GetExcelColName(i)}2), char(34)";
                else
                    answer += $"{GetExcelColName(i)}2";
                if (i != columnCount)
                    answer += ", \",\", ";
                else
                    answer += ", ";
            }
            answer += $"\")\" )";


            Console.WriteLine(answer);



        }
    }
}
