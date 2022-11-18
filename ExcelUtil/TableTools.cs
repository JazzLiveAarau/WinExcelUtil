using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace ExcelUtil
{
    /// <summary>Functions for modification of tables</summary>
    static public class TableTools
    {
        /// <summary>Remove table column defined by the column index</summary>
        static public bool RemoveColumn(ref Table io_table, int i_column_index, out string o_error)
        {
            o_error = "";

            for (int i_row = 0; i_row < io_table.NumberRows; i_row++)
            {
                Row current_row = io_table.GetRow(i_row, out o_error);
                if (o_error != "") return false;

                 if (!current_row.RemoveField(i_column_index, out o_error)) return false;
            }

            io_table.NumberColumns = io_table.NumberColumns - 1;

            return true;
        }

        /// <summary>Remove table column defined by the column header</summary>
        static public bool RemoveColumn(ref Table io_table, string i_column_header, out string o_error)
        {
            o_error = "";

            Row first_row = io_table.GetRow(0, out o_error);

            int column_index = Table.GetColumnIndex(first_row, i_column_header);

            if (column_index < 0)
            {
                o_error = "TableTools.RemoveColumn There is no column with header " + i_column_header;
                return false;
            }

            return RemoveColumn(ref io_table, column_index, out o_error);
        }

        /// <summary>Create a column from an array of strings that are the column field values</summary>
        static public bool CreateColumn(string[] i_fields_as_strings, out Column o_column, out string o_error)
        {
            o_error = "";

            o_column = new Column();

            for (int i_field = 0; i_field < i_fields_as_strings.Length; i_field++)
            {
                string field_value = i_fields_as_strings[i_field];

                Field current_field = new Field(field_value);

                o_column.AppendField(current_field);
            }

            return true;
        }

        /// <summary>Get a table column defined by the column index.</summary>
        /// <param name="i_table">Input table</param>
        /// <param name="i_column_index">Column index</param>
        /// <param name="o_column">Output column</param>
        /// <param name="o_field_header">Output field header if existing. Returned value is null if not</param>
        /// <param name="o_error">Error message if function has failed</param>
        /// <returns></returns>
        static public bool GetColumn(Table i_table, int i_column_index, out Column o_column, out FieldHeader o_field_header, out string o_error)
        {
            o_column = null;
            o_field_header = null;
            o_error = "";

            if (i_column_index < 0 || i_column_index >= i_table.NumberColumns)
            {
                o_error = "Column index " + i_column_index.ToString() + " is not between 0 and " + i_table.NumberColumns.ToString();
                return false;
            }

            o_column = new Column();

            for (int i_field = 0; i_field < i_table.NumberRows; i_field++)
            {
                Row current_row = i_table.GetRow(i_field, out o_error);
                if (o_error != "") return false;


                Field row_field = current_row.GetField(i_column_index, out o_error);
                if (o_error != "") return false;

                // A new field object must be created
                Field current_field = new Field(row_field.FieldValue);

                o_column.AppendField(current_field);
            }

            return true;
        }

        /// <summary>Move a column</summary>
        static public bool MoveColumn(ref Table io_table, string i_column_header, string i_column_move_before, out string o_error)
        {
            o_error = "";

            if (io_table.HasHeaderData())
            {
                o_error = "TableTools.MoveColumn The input table has header data";
                return false;
            }

            if (i_column_header == i_column_move_before)
            {
                o_error = "TableTools.MoveColumn i_column_header= i_column_move_before";
                return false;
            }

            Column move_column = null;
            FieldHeader field_header;
            if (!GetColumn(io_table, i_column_header, out move_column, out field_header, out o_error)) return false;

            if (!RemoveColumn(ref io_table, i_column_header, out o_error)) return false;

            if (!InsertColumn(ref io_table, i_column_move_before, move_column, out o_error)) return false;

            return true;
        }

        /// <summary>Get a table column defined by the column header. TODO FieldHeader</summary>
        /// <param name="i_table">Input table</param>
        /// <param name="i_column_index">Column index</param>
        /// <param name="o_column">Output column</param>
        /// <param name="o_field_header">Output field header if existing. Returned value is null if not</param>
        /// <param name="o_error">Error message if function has failed</param>
        /// <returns></returns>
        static public bool GetColumn(Table i_table, string i_column_header, out Column o_column, out FieldHeader o_field_header, out string o_error)
        {
            o_column = null;
            o_field_header = null;
            o_error = "";

            Row first_row = i_table.GetRow(0, out o_error);
            if (o_error != "") return false;

            int column_index = Table.GetColumnIndex(first_row, i_column_header);

            if (column_index < 0)
            {
                o_error = "TableTools.InsertColumn There is no column with header " + i_column_header;
                return false;
            }

            return GetColumn(i_table, column_index, out o_column, out o_field_header, out o_error);
        }

        /// <summary>Get table columns defined by a search string. TODO FieldHeaders</summary>
        static public bool GetColumns(Table i_table, string i_column_search, out Column[] o_columns, out FieldHeader[] o_field_headers, out string o_error)
        {
            o_error = "";

            ArrayList array_list_columns = new ArrayList();
            ArrayList array_list_headers = new ArrayList();
            o_columns = (Column[])array_list_columns.ToArray(typeof(Column));
            o_field_headers = (FieldHeader[])array_list_headers.ToArray(typeof(FieldHeader));


            Row first_row = i_table.GetRow(0, out o_error);
            if (o_error != "") return false;

            for (int i_field = 0; i_field < first_row.NumberColumns; i_field++)
            {
                Field current_field = first_row.GetField(i_field, out o_error);
                if (o_error != "") return false;

                if (current_field.FieldValue.Contains(i_column_search))
                {
                    Column search_column;
                    FieldHeader field_header;

                    if (!GetColumn(i_table, i_field, out search_column, out field_header, out o_error)) return false;

                    array_list_columns.Add(search_column);
                    // TODO field_header
                }
            }

            o_columns = (Column[])array_list_columns.ToArray(typeof(Column));
            o_field_headers = (FieldHeader[])array_list_headers.ToArray(typeof(FieldHeader));

            return true;
        }

        /// <summary>Insert table column defined by the column index. Input is column object and a field header object</summary>
        static public bool InsertColumn(ref Table io_table, int i_column_index, Column i_column, FieldHeader i_field_header, out string o_error)
        {
            o_error = "";

            if (!io_table.HasHeaderData())
            {
                o_error = "TableTools.InsertColumn The input table has no header data";
                return false;
            }

            int number_rows = io_table.NumberRows;
            int number_fields = i_column.NumberFields;

            if (number_fields != number_rows)
            {
                o_error = "TableTools.InsertColumn Number of rows " + number_rows.ToString() + " not equal to the number of fields " + number_fields.ToString();
                return true;
            }

            for (int i_row = 0; i_row < io_table.NumberRows; i_row++)
            {
                Row current_row = io_table.GetRow(i_row, out o_error);
                if (o_error != "") return false;

                Field current_field = i_column.GetField(i_row, out o_error);
                if (o_error != "") return false;

                if (!current_row.InsertField(i_column_index, current_field, out o_error)) return false;
            }

            o_error = "TableTools.InsertColumn This function is not yet implemented";
            return false; // QQQQQQ
        }

        /// <summary>Insert table column defined by the column header. Input is column object and a field header object</summary>
        static public bool InsertColumn(ref Table io_table, string i_column_header, Column i_column, FieldHeader i_field_header, out string o_error)
        {
            o_error = "";

            Row first_row = io_table.GetRow(0, out o_error);

            int column_index = Table.GetColumnIndex(first_row, i_column_header);

            if (column_index < 0)
            {
                o_error = "TableTools.InsertColumn There is no column with header " + i_column_header;
                return false;
            }

            return InsertColumn(ref io_table, column_index, i_column, i_field_header, out o_error);
        }

        /// <summary>Insert table column defined by the column index. Input is column object (and no field header object)</summary>
        static public bool InsertColumn(ref Table io_table, int i_column_index, Column i_column, out string o_error)
        {
            o_error = "";

            if (io_table.HasHeaderData())
            {
                o_error = "TableTools.InsertColumn The input table has header data";
                return false;
            }

            int number_rows = io_table.NumberRows;
            int number_fields = i_column.NumberFields;

            if (number_fields != number_rows)
            {
                o_error = "TableTools.InsertColumn Number of rows " + number_rows.ToString() + " not equal to the number of fields " + number_fields.ToString();
                return true;
            }

            for (int i_row = 0; i_row < io_table.NumberRows; i_row++)
            {
                Row current_row = io_table.GetRow(i_row, out o_error);
                if (o_error != "") return false;

                Field current_field = i_column.GetField(i_row, out o_error);
                if (o_error != "") return false;

                if (!current_row.InsertField(i_column_index, current_field, out o_error)) return false;
            }

            io_table.NumberColumns = io_table.NumberColumns + 1;

            return true;
        }

        /// <summary>Insert table column defined by the column header. Input is column object (and no field header object)</summary>
        static public bool InsertColumn(ref Table io_table, string i_column_header, Column i_column, out string o_error)
        {
            o_error = "";

            Row first_row = io_table.GetRow(0, out o_error);

            int column_index = Table.GetColumnIndex(first_row, i_column_header);

            if (column_index < 0)
            {
                o_error = "TableTools.InsertColumn There is no column with header " + i_column_header;
                return false;
            }

            return InsertColumn(ref io_table, column_index, i_column, out o_error);
        }


        /// <summary>Change name of table column defined by the column index.</summary>
        static public bool ChangeColumnName(ref Table io_table, int i_column_index, string i_new_name, out string o_error)
        {
            o_error = "";

            Row first_row = io_table.GetRow(0, out o_error);
            if (o_error != "") return false;

            if (!first_row.SetFieldValue(i_column_index, i_new_name, out o_error)) return false;

            return true;
        }

        /// <summary>Change name of table column defined by the column string.</summary>
        static public bool ChangeColumnName(ref Table io_table, string i_column_header, string i_new_name, out string o_error)
        {
            o_error = "";

            Row first_row = io_table.GetRow(0, out o_error);

            int column_index = Table.GetColumnIndex(first_row, i_column_header);

            if (column_index < 0)
            {
                o_error = "TableTools.ChangeColumnName There is no column with header " + i_column_header;
                return false;
            }

            return ChangeColumnName(ref io_table, column_index, i_new_name, out o_error);
        }

        /// <summary>Change name of column.</summary>
        static public bool ChangeColumnName(ref Column io_column, string i_new_name, out string o_error)
        {
            o_error = "";

            Field first_field = io_column.GetField(0, out o_error);
            if (o_error != "") return false;

            first_field.FieldValue = i_new_name;

            return true;
        }
    }
}
