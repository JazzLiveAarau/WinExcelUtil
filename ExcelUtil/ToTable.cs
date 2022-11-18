using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using DetectEncoding;
using System.Collections;

namespace ExcelUtil
{
    /// <summary>Has functions that convert "table" files in different formats (CSV, XSL, XSLX, XML, ..) to a Table</summary>
    public class ToTable
    {
        /// <summary>Delimiters for csv files</summary>
        public static string[] m_csv_delimiters = { ";", ",", ":" };

        /// <summary>Convert the .csv file to a table.</summary>
        static public bool CsvToTable(string i_file_csv, ref Table io_table, ref Encoding io_file_encoding, out string o_error)
        {
            o_error = "";

            if (!File.Exists(i_file_csv))
            {
                o_error = "ToTable.CsvToTable There is no file " + i_file_csv;
                return false;
            }

            string csv_delimiter = "";
            return CsvToTable(i_file_csv, ref io_table, out csv_delimiter, ref io_file_encoding, out o_error);
        }

        /// <summary>Convert the .csv file to a table.</summary>
        static public bool CsvToTable(string i_file_csv, ref Table io_table, out string o_error)
        {
            o_error = "";

            if (!File.Exists(i_file_csv))
            {
                o_error = "ToTable.CsvToTable There is no file " + i_file_csv;
                return false;
            }

            string csv_delimiter = "";
            Encoding file_encoding = null;
            return CsvToTable(i_file_csv, ref io_table, out csv_delimiter, ref file_encoding, out o_error);
        }

        /// <summary>Convert the .csv file to a table. Output is also the delimiter.</summary>
        /// <param name="i_file_csv">Input file name</param>
        /// <param name="io_table">Output table</param>
        /// <param name="o_csv_delimiter">Output delimiter that is used in the .csv file</param>
        /// <param name="io_file_encoding">Detected encoding is returned if input is null. If encoding is set it will be used.</param>
        /// <param name="o_error">Error message if function has failed</param>
        static public bool CsvToTable(string i_file_csv, ref Table io_table, out string o_csv_delimiter, ref Encoding io_file_encoding, out string o_error)
        {
            o_error = "";
            o_csv_delimiter = "";

            if (!File.Exists(i_file_csv))
            {
                o_error = "ToTable.CsvToTable There is no file " + i_file_csv;
                return false;
            }

            int number_fields = -1;

            int n_rows = 0;

            if (null == io_file_encoding)
            {
                if (!_DetectEncoding(i_file_csv, ref io_file_encoding, out o_error)) return false;
            }

            try
            {
                using (FileStream file_stream = new FileStream(i_file_csv, FileMode.Open, FileAccess.Read, FileShare.Read))
                // Without System.Text.Encoding.UTF8 there are problems with ä ö ü. With Encoding.Default it worked in some computers
                // Alternatives Encoding.Default, Encoding.UTF8, Encoding.Unicode, Encoding.UTF32, Encoding.UTF7
                // using (StreamReader stream_reader = new StreamReader(file_stream, System.Text.Encoding.Default))
                using (StreamReader stream_reader = new StreamReader(file_stream, io_file_encoding))
                {
                    while (stream_reader.Peek() >= 0)
                    {
                        string current_row = stream_reader.ReadLine();
                        n_rows = n_rows + 1;
                        if (1 == n_rows)
                        {
                            o_csv_delimiter = _CsvDelimiter(current_row, out number_fields);
                            if ("" == o_csv_delimiter)
                            {
                                o_error = "No delimiter found in the first row";
                                return false;
                            }
                        } // 1 == n_rows

                        
                        if (current_row.Trim() == "")
                        {
                            // A line with only spaces. Probably the last line(s)
                            n_rows = n_rows - 1;
                            break;
                        }

                        Row new_row;
                        if (!_CsvCreateRow(current_row, o_csv_delimiter, number_fields, out new_row, out o_error)) return false;

                        if (!io_table.AddRow(new_row, out o_error)) return false;

                    } // while
                }
            }


            catch (FileNotFoundException) {o_error = "File not found"; return false; }
            catch (DirectoryNotFoundException) { o_error = "Directory not found"; return false; }
            catch (InvalidOperationException) { o_error = "Invalid operation"; return false; }
            catch (InvalidCastException) { o_error = "invalid cast"; return false; } 
            catch (Exception e)
            {
                o_error = " Unhandled Exception " + e.GetType() + " occurred at " + DateTime.Now + "!";
                return false;
            }

            io_table.NumberRows = n_rows;
            io_table.NumberColumns = number_fields;

            return true;
        } // CsvToTable

        /// <summary>Returns one row in the .csv file as a Row object</summary>
        static private bool _CsvCreateRow(string i_row, string csv_delimiter, int i_number_fields, out Row o_row, out string o_error)
        {
            o_error = "";
            o_row = new Row();

            int n_delimiters = 0;

            int n_check_fields = 0;

            string current_field = "";

            for (int i_char = 0; i_char < i_row.Length; i_char++)
            {
                string c_char = i_row.Substring(i_char, 1);

                if (csv_delimiter != c_char && i_char == i_row.Length - 1)
                {
                    current_field = current_field + c_char;
                }

                if (csv_delimiter == c_char)
                {
                    Field new_field = new Field(current_field);
                    o_row.AddField(new_field);
                    current_field = "";

                    n_delimiters = n_delimiters + 1;

                    n_check_fields = n_check_fields + 1;
                }
                else if (i_row.Length - 1 == i_char)
                {
                    Field new_field = new Field(current_field);
                    o_row.AddField(new_field);
                    current_field = "";

                    n_check_fields = n_check_fields + 1;
                }
                else
                {
                    current_field = current_field + c_char;
                }
            }

            if (n_delimiters + 1 != i_number_fields)
            {
                o_error = "Number of fields is not the same as for the first row \n Row: " + i_row;
                return false;
            }

            if (n_check_fields != i_number_fields)
            {
                o_error = "Number of records is not as for the first row \n Row: " + i_row;
                return false;
            }

            return true;
        }

        /// <summary>Determines and returns the delimiter between the fields in a row from a .csv file</summary>
        static private string _CsvDelimiter(string i_row, out int o_number_fields)
        {
            string ret_delimiter = "";
            o_number_fields = -1;

            string[] all_delimiters = m_csv_delimiters;
            
            int[] n_delimiters;
            ArrayList n_array_list_delimiters = new ArrayList();
            for (int i_init = 0; i_init < all_delimiters.Length; i_init++)
            {
                n_array_list_delimiters.Add(0);
            }
            n_delimiters = (int[])n_array_list_delimiters.ToArray(typeof(int));


            for (int i_delimiter = 0; i_delimiter < all_delimiters.Length; i_delimiter++)
            {
                string c_delimiter = all_delimiters[i_delimiter];
                

                for (int i_char = 0; i_char < i_row.Length; i_char++)
                {
                    string c_char = i_row.Substring(i_char, 1);

                    if (c_delimiter == c_char)
                    {
                        n_delimiters[i_delimiter] = n_delimiters[i_delimiter] + 1;
                    }
                }

            }

            int max_index = -1;
            int max_number = -1;
            for (int i_delimiter = 0; i_delimiter < all_delimiters.Length; i_delimiter++)
            {
                if (n_delimiters[i_delimiter] > max_number)
                {
                    max_number = n_delimiters[i_delimiter];
                    max_index = i_delimiter;
                }
            }

            if (max_index >= 0)
            {
                ret_delimiter = all_delimiters[max_index];
                o_number_fields = max_number + 1;
            }

            return ret_delimiter;

        } // _CsvDelimiter

        /// <summary>Detect the encoding of a file</summary>
        static private bool _DetectEncoding(string i_file_csv, ref Encoding o_encoding, out string o_error)
        {
            o_error = "";
            o_encoding = null;

            if (!File.Exists(i_file_csv))
            {
                o_error = "ToTable._DetectEncoding There is no file " + i_file_csv;
                return false;
            }

            // create a StreamReader with the guessed best encoding
            string file_content_str = "";
            using (StreamReader stream_reader = EncodingTools.OpenTextFile(i_file_csv))
            {
                file_content_str = stream_reader.ReadToEnd();
            }

            o_encoding = EncodingTools.GetMostEfficientEncoding(file_content_str);


            return true;
        }

    }
}
