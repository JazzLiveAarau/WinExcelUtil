using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace ExcelUtil
{
    /// <summary>Functions for finding rows with a given string</summary>
    static public class TableSearch
    {
        /// <summary>Get rows with a given string. Search columns are defined with an array of column indices</summary>
        static public bool GetRows(Table i_table, int[] i_column_indices, string i_search_string, out int[] row_indices, out string o_error)
        {
            o_error = "";

            ArrayList array_list_row_indices = new ArrayList();
            row_indices = (int[])array_list_row_indices.ToArray(typeof(int));

            string search_string_upper = i_search_string.ToUpper();

            // Note index start is one (1). No search in the header row
            for (int i_row = 1; i_row < i_table.NumberRows; i_row++)
            {
                Row current_row = i_table.GetRow(i_row, out o_error);
                if (o_error != "") return false;

                for (int i_column = 0; i_column < i_column_indices.Length; i_column++)
                {
                    int column_index = i_column_indices[i_column];

                    Field current_field = current_row.GetField(column_index, out o_error);
                    if (o_error != "") return false;

                    string current_value = current_field.FieldValue;
                    string current_value_upper = current_value.ToUpper();

                    if (current_value_upper.Contains(search_string_upper))
                    {
                        array_list_row_indices.Add(i_row);
                        // Break loop i_column after one hit. E-Mail address and name often same
                        break;
                    }
                }
            }

            row_indices = (int[])array_list_row_indices.ToArray(typeof(int));

            return true;
        }

        /// <summary>Get rows with a given string. Search column is defined with a column index</summary>
        static public bool GetRows(Table i_table, int i_column_index, string i_search_string, out int[] row_indices, out string o_error)
        {
            row_indices = null;
            o_error = "";

            int[] column_indices = new int[1];
            column_indices[0] = i_column_index;

            return GetRows(i_table, column_indices, i_search_string, out row_indices, out o_error);
        }

        /// <summary>Get rows with a given string. Search column is defined by a column header</summary>
        static public bool GetRows(Table i_table, string i_column_header, string i_search_string, out int[] row_indices, out string o_error)
        {
            row_indices = null;
            o_error = "";

            Row first_row = i_table.GetRow(0, out o_error);

            int column_index = Table.GetColumnIndex(first_row, i_column_header);

            if (column_index < 0)
            {
                o_error = "TableSearch.GetRows There is no column with header " + i_column_header;
                return false;
            }

            return GetRows(i_table, column_index, i_search_string, out row_indices, out o_error);
        }

        /// <summary>Get rows with a given string. Search columns are defined by an array column headers</summary>
        static public bool GetRows(Table i_table, string[] i_column_headers, string i_search_string, out int[] row_indices, out string o_error)
        {
            row_indices = null;
            o_error = "";

            Row first_row = i_table.GetRow(0, out o_error);

            int[] column_indices = new int[i_column_headers.Length];

            for (int i_column = 0; i_column < i_column_headers.Length; i_column++)
            {
                string column_header = i_column_headers[i_column];

                int column_index = Table.GetColumnIndex(first_row, column_header);

                if (column_index < 0)
                {
                    o_error = "TableSearch.GetRows There is no column with header " + column_header;
                    return false;
                }

                column_indices[i_column] = column_index;
            }

            return GetRows(i_table, column_indices, i_search_string, out row_indices, out o_error);
        }

    }
}
