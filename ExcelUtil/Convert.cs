using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;


namespace ExcelUtil
{
    /// <summary>Has functions that convert "table" files to other "table" files</summary>
    public class Convert
    {
        /// <summary>Converts a .csv file to a .xslx file</summary>
        /// <param name="i_file_csv">Input .csv file name with path</param>
        /// <param name="i_file_xsl">Input .xlsx file name with path</param>
        /// <param name="o_error">Error message if function fails</param>
        /// <returns></returns>
        static public bool CsvToXlsx(string i_file_csv, string i_file_xsl, out string o_error)
        {
            o_error = "";

            if (!_CsvToXlsxCheckInput(i_file_csv, i_file_xsl, out o_error)) return false;

            Table table_csv = new Table("Table created from a CSV file");

            if (!ToTable.CsvToTable(i_file_csv, ref table_csv, out o_error)) return false;

            if (!FromTable.TableToXlsx(table_csv, i_file_xsl, out o_error)) return false;

            return true;
        }

        /// <summary>Checks input data for function CsvToXlsx</summary>
        static public bool _CsvToXlsxCheckInput(string i_file_csv, string i_file_xsl, out string o_error)
        {
            o_error = "";

            if (!File.Exists(i_file_csv))
            {
                o_error = "No file " + i_file_csv;
                return false;
            }

            // Add check of name and path of i_file_xsl

            return true;
        }
    }
}
