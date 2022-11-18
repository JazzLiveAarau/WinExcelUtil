using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

// Add Reference -> COM -> Microsoft Excel 14.0 Object Library
// Made reference to 12.0 in Anna-Lisas computer
using Excel = Microsoft.Office.Interop.Excel;
// In References displayed as Microsoft.Office.core, Microsoft.Office.Interop.Excel and VBIDE

namespace ExcelUtil
{
    /// <summary>Has functions that convert a Table to "table" file in different formats (CSV, XSL, XSLX, XML, ..)</summary>
    public class FromTable
    {
        /// <summary>Convert a table to a .csv file</summary>
        /// <param name="i_table">Input table</param>
        /// <param name="i_file_csv">Full name (with extension) of the output .csv file</param>
        /// <param name="i_init">The delimiter that separates the fields</param>
        /// <param name="i_file_encoding">Encoding for output file. If null default encoding will be used</param>
        /// <param name="o_error">Error message if function fails</param>
        /// <returns></returns>
        static public bool TableToCsv(Table i_table, string i_file_csv, string i_delimiter, Encoding i_file_encoding, out string o_error)
        {
            o_error = "";

            if (!_CheckInputTable(i_table, out o_error)) return false;

            bool b_delimiter_allowed = false;
            for (int i_delim = 0; i_delim < ToTable.m_csv_delimiters.Length; i_delim++)
            {
                string allowed_delimiter = ToTable.m_csv_delimiters[i_delim];

                if (i_delimiter == allowed_delimiter)
                {
                    b_delimiter_allowed = true;
                    break;
                }
            }

            if (!b_delimiter_allowed)
            {
                o_error = "FromTable.TableToCsv Delimiter " + i_delimiter + " is not allowed";
                return false;
            }

            Encoding file_encoding = i_file_encoding;
            if (null == file_encoding)
                file_encoding = System.Text.Encoding.Default;


            try
            {
                using (FileStream fileStream = new FileStream(i_file_csv, FileMode.Create))
                // Without System.Text.Encoding.Default there are problems with ä ö ü
                // using (StreamWriter stream_writer = new StreamWriter(file_stream, System.Text.Encoding.Default))
                using (StreamWriter stream_writer = new StreamWriter(fileStream, file_encoding))
                {
                    for (int i_row = 0; i_row < i_table.NumberRows; i_row++)
                    {
                        Row current_row = i_table.GetRow(i_row, out o_error);
                        if (o_error != "") return false;

                        for (int i_field = 0; i_field < current_row.NumberColumns; i_field++)
                        {
                            Field current_field = current_row.GetField(i_field, out o_error);
                            if (o_error != "") return false;

                            string current_field_value = current_field.FieldValue;

                            if (i_field < current_row.NumberColumns - 1)
                            {
                                stream_writer.Write(current_field_value + i_delimiter);
                            }
                            else
                            {
                                stream_writer.Write(current_field_value);
                            }
                        }
                        stream_writer.Write("\n");
                    }

                    stream_writer.Close();
                }
            }

            catch (FileNotFoundException) { o_error = "File not found"; return false; }
            catch (DirectoryNotFoundException) { o_error = "Directory not found"; return false; }
            catch (InvalidOperationException) { o_error = "Invalid operation"; return false; }
            catch (InvalidCastException) { o_error = "invalid cast"; return false; }
            catch (Exception e)
            {
                o_error = " Unhandled Exception " + e.GetType() + " occurred at " + DateTime.Now + "!";
                return false;
            }

            
            return true;
        }

        /// <summary>Convert a table to a .xlsx file</summary>
        static public bool TableToXlsx(Table i_table, string i_file_xlsx, out string o_error)
        {
            o_error = "";

            if (!_CheckInputTable(i_table, out o_error)) return false;

            Excel.Application excel_application = new Excel.Application();

            Excel.Workbook excel_workbook = null;

            try
            {
                excel_workbook = excel_application.Workbooks.Add(Excel.XlWBATemplate.xlWBATWorksheet);
            }
            catch
            {
                o_error = "FromTable.TableToXlsx: Open of workbook failed";
                excel_workbook = null;
                return false;
            }

            Excel.Worksheet excel_worksheet = null;

            int n_sheet = excel_workbook.Worksheets.Count;
            if (n_sheet == 0 )
            {
                o_error = "FromTable.TableToXlsx: Number of work sheets = 0";
                return false;
            }

            try
            {
                excel_worksheet = (Excel.Worksheet)excel_workbook.Worksheets.get_Item(1);
            }
            catch
            {
                o_error = "FromTable.TableToXlsx: Getting the worksheet failed";
                excel_application.Quit();
                excel_worksheet = null;
            }

            for (int i_row = 0; i_row < i_table.NumberRows; i_row++)
            {
                Row current_row = i_table.GetRow(i_row, out o_error);
                if (o_error != "") return false;

                for (int i_field = 0; i_field < current_row.NumberColumns; i_field++)
                {
                    Field current_field = current_row.GetField(i_field, out o_error);
                    if (o_error != "") return false;

                    string current_field_value = current_field.FieldValue;

                    excel_worksheet.Cells[i_row + 1, i_field + 1] = current_field_value;
                }
            }


            if (File.Exists(i_file_xlsx))
            {
                File.Delete(i_file_xlsx);
            }

            if (excel_worksheet != null)
            {
                try
                {
                    excel_workbook.SaveAs(i_file_xlsx);
                }
                catch
                {
                    o_error = "FromTable.TableToXlsx: Workbook SaveAs failed";
                }

                object misValue = System.Reflection.Missing.Value;

                try
                {
                    bool b_save_changes = true;
                    excel_workbook.Close(b_save_changes, i_file_xlsx, misValue);
                }
                catch
                {
                    o_error = "FromTable.TableToXlsx: Save and close failed";
                }
            }


            excel_application.Quit();

            /*            
System.Text.Encoding.Default
            */

            return true;
        }

        /// <summary>Convert a table to an .xml file</summary>
        /// <param name="i_table">Input table</param>
        /// <param name="i_file_xml">Full name (with extension) of the output .csv file</param>
        /// <param name="i_xml_tag_person">For instance Supporter oder Sponsor</param>
        /// <param name="i_file_encoding">Encoding for output file. If null default encoding will be used</param>
        /// <param name="o_error">Error message if function fails</param>
        /// <returns>True for succesful file creation</returns>
        static public bool TableToXml(Table i_table, string i_file_xml, string i_xml_tag_person, Encoding i_file_encoding, out string o_error)
        {
            o_error = "";

            if (!_CheckInputTable(i_table, out o_error)) return false;

            if (i_xml_tag_person.Length < 3)
            {
                o_error = "FromTable.TableToXml Minimum length is three (3) for the person XML tag (" + i_xml_tag_person + ")";

                return false;
            }

            string[] xml_tags_array = _GetXmlTags(i_table, out o_error);
            if (xml_tags_array == null || o_error.Length > 0)
            {
                return false;
            }

            Encoding file_encoding = i_file_encoding;
            if (null == file_encoding)
                file_encoding = System.Text.Encoding.Default;

            try
            {
                using (FileStream fileStream = new FileStream(i_file_xml, FileMode.Create))
 
                using (StreamWriter stream_writer = new StreamWriter(fileStream, file_encoding))
                {
                    string start_tag = "<AddressesData>";

                    stream_writer.Write(start_tag);

                    stream_writer.Write("\n");

                    for (int i_row = 1; i_row < i_table.NumberRows; i_row++)
                    {
                        Row current_row = i_table.GetRow(i_row, out o_error);
                        if (o_error != "") return false;

                        string start_tag_person = "    <" + i_xml_tag_person + ">";

                        stream_writer.Write(start_tag_person);

                        stream_writer.Write("\n");

                        for (int i_field = 0; i_field < current_row.NumberColumns; i_field++)
                        {
                            Field current_field = current_row.GetField(i_field, out o_error);
                            if (o_error != "") return false;

                            string current_field_value = current_field.FieldValue;

                            string current_xml_value =  _GetXmlValue(current_field_value);

                            string current_xml_tag = xml_tags_array[i_field];

                            string row_xml = "        <" + current_xml_tag + ">" + current_xml_value + "</" + current_xml_tag + ">";

                            stream_writer.Write(row_xml);

                            stream_writer.Write("\n");

                        } // i_field

                        string end_tag_person = "    </" + i_xml_tag_person + ">";

                        stream_writer.Write(end_tag_person);

                        stream_writer.Write("\n");

                    } // i_row

                    string end_tag = "</AddressesData>";

                    stream_writer.Write(end_tag);

                    stream_writer.Write("\n");

                    stream_writer.Close();
                }
            }

            catch (FileNotFoundException) { o_error = "File not found"; return false; }
            catch (DirectoryNotFoundException) { o_error = "Directory not found"; return false; }
            catch (InvalidOperationException) { o_error = "Invalid operation"; return false; }
            catch (InvalidCastException) { o_error = "invalid cast"; return false; }
            catch (Exception e)
            {
                o_error = " Unhandled Exception " + e.GetType() + " occurred at " + DateTime.Now + "!";
                return false;
            }


            return true;

        } // TableToXml

        /// <summary>
        /// Not allowed characters for an XML value is &, < and <. These are removed.
        /// Empty value is also not allowed. For this case NotYetSetNodeValue is returned
        /// </summary>
        /// <param name="i_field_value">Input field value from the table</param>
        /// <returns>String (reduced) with allowed XML characters</returns>
        static private string _GetXmlValue(string i_field_value)
        {
            string ret_xml_value = "";

            string field_value_trim = i_field_value.Trim();

            if (field_value_trim.Length == 0)
            {
                return "NotYetSetNodeValue";
            }

            for (int index_char=0; index_char < field_value_trim.Length; index_char++)
            {
                string current_char = field_value_trim.Substring(index_char, 1);

                if (!current_char.Equals("&") && !current_char.Equals("<") && !current_char.Equals("<"))
                {
                    ret_xml_value = ret_xml_value + current_char;
                }
            }

            ret_xml_value = ret_xml_value.Trim();

            if (ret_xml_value.Length == 0)
            {
                return "NotYetSetNodeValue";
            }

            return ret_xml_value;

        } // _GetXmlValue


        /// <summary>First row of the table defines the tags for the XML file</summary>
        /// <param name="i_table">Input table</param>
        /// <param name="o_error">Error description if function fails</param>
        /// <returns>Array of XML tags</returns>
        static private string[] _GetXmlTags(Table i_table, out string o_error)
        {
            o_error = "";

            int row_number = 0;
            Row start_row = i_table.GetRow(row_number, out o_error);
            if (o_error != "") return null;

            int number_tags = start_row.NumberColumns;

            string[]ret_tags_array = new string[number_tags];

            for (int i_field = 0; i_field < start_row.NumberColumns; i_field++)
            {
                Field current_field = start_row.GetField(i_field, out o_error);
                if (o_error != "") return null;

                string current_field_value = current_field.FieldValue;

                ret_tags_array[i_field] = current_field_value;
            }

            return ret_tags_array;

        } // _GetXmlTags

        /// <summary>Check Table data</summary>
        static private bool _CheckInputTable(Table i_table, out string o_error)
        {
            o_error = "";

            if (null == i_table)
            {
                o_error = "_CheckInputTable Input table is null";
                return false;
            }

            if (i_table.NumberRows < 2)
            {
                o_error = "_CheckInputTable Number of rows is less than two";
                return false;
            }

            if (i_table.NumberColumns < 1)
            {
                o_error = "_CheckInputTable Number of columns is less than one";
                return false;
            }

            return true;
        }
    }
}
