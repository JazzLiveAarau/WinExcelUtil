using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelUtil
{
    /// <summary>Functions for sorting of table rows</summary>
    static public class TableSort
    {
        /// <summary>Sort alphabetically for a given field defined by the column index</summary>
        static public bool SortField(ref Table io_table, int i_column_index, out string o_error)
        {
            o_error = "";

            Column sort_column; 
            FieldHeader field_header;
            if (!TableTools.GetColumn(io_table, i_column_index, out sort_column, out field_header, out o_error)) return false;

            string[] sort_array = null;
            int[] sort_keys = null;

            if (!_SortArrays(sort_column, out sort_array, out sort_keys, out o_error)) return false;

            Array.Sort(sort_array, sort_keys);

            int add_index = 0;

            for (int i_sort = 0; i_sort < sort_keys.Length; i_sort++)
            {
                int index_row = sort_keys[i_sort] + add_index;

                int index_to_row = i_sort + 1;

                if (!io_table.MoveRow(index_row, index_to_row, out o_error)) return false;

                if (!_ModifyKeysArray(ref sort_keys, index_row, index_to_row, out o_error)) return false;
            }


            return true;
        }

        /// <summary>Modify the keys array when a row has been moved</summary>
        static private bool _ModifyKeysArray(ref int[] io_sort_keys, int i_index_moved_row, int i_index_to_row, out string o_error)
        {
            o_error = "";

            for (int i_modify = i_index_to_row; i_modify < io_sort_keys.Length; i_modify++)
            {
                if (io_sort_keys[i_modify] < i_index_moved_row)
                {
                    io_sort_keys[i_modify] = io_sort_keys[i_modify] + 1;
                }
            }

            return true;
        }

        /// <summary>Get array to sort and the corresponding keys array with row numbers</summary>
        static private bool _SortArrays(Column i_sort_column, out string[] o_sort_array, out int[] o_sort_keys, out string o_error)
        {
            o_error = "";

            // Note NumberFields-1: The first field value (header) shall not be part of the sorting
            o_sort_array = new string[i_sort_column.NumberFields - 1];
            o_sort_keys = new int[i_sort_column.NumberFields - 1];

            // Note start index 1. First field shall not be sorted.
            for (int i_field = 1; i_field < i_sort_column.NumberFields; i_field++)
            {
                Field current_field = i_sort_column.GetField(i_field, out o_error);
                if (o_error != "") return false;

                // Note index i_field - 1
                o_sort_array[i_field - 1] = current_field.FieldValue;
                o_sort_keys[i_field - 1] = i_field;
            }

            return true;
        }

        /// <summary>Sort alphabetically for a given field defined by the column header</summary>
        static public bool SortField(ref Table io_table, string i_column_header, out string o_error)
        {
            o_error = "";

            Row first_row = io_table.GetRow(0, out o_error);

            int column_index = Table.GetColumnIndex(first_row, i_column_header);

            if (column_index < 0)
            {
                o_error = "TableSort.SortField There is no column with header " + i_column_header;
                return false;
            }

            return SortField(ref io_table, column_index, out o_error);
        }
    }
}
