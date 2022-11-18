using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Collections;

namespace ExcelUtil
{
    /// <summary>Holds the data for a table. Has functions to set and retrieve data</summary>
    public class Table
    {
        /// <summary>Name of the table</summary>
        private string m_name = "";

        /// <summary>Header data for the rows (columns)</summary>
        private RowHeader m_row_header = null;

        /// <summary>List of rows</summary>
        private ArrayList m_array_rows = new ArrayList();

        /// <summary>All rows in the table</summary>
        private Row[] m_rows;

        /// <summary>Number of rows in the table</summary>
        private int m_number_rows = 0;

        /// <summary>Number of colums in the table</summary>
        private int m_number_colums = 0;

        /// <summary>Constructor with name as input</summary>
        public Table(string i_name)
        {
            m_name = i_name;
        }

        /// <summary>Constructor with name and row (column) header data as input</summary>
        public Table(string i_name, RowHeader i_row_header)
        {
            m_name = i_name;

            m_row_header = i_row_header;
        }

        /// <summary>Sets the header row</summary>
        public void SetRowHeader(RowHeader i_row_header)
        {
            if (null == i_row_header)
                return; // Programming error

            m_row_header = i_row_header;
        }

        /// <summary>Returns the header row</summary>
        public RowHeader GetRowHeader()
        {
            return m_row_header;
        }

        /// <summary>Returns true if the table has header data</summary>
        public bool HasHeaderData()
        {
            if (null == m_row_header)
            {
                return false;
            }
            else
            {
                return true;
            }
        }

        /// <summary>Sets a field value that is a string. Field is defined by a row index and a field index</summary>
        public bool SetFieldString(int i_index_row, int i_index_field, string i_field_value, out string o_error)
        {
            o_error = "";

            Row current_row = GetRow(i_index_row, out o_error);
            if ("" != o_error) return false;

            Field current_field = current_row.GetField(i_index_field, out o_error);
            if ("" != o_error) return false;

            current_field.FieldValue = i_field_value;

            return true;
        }

        /// <summary>Sets a field value that is a string. Field is defined by a row index andand the column name (field value in the first row)</summary>
        public bool SetFieldString(int i_index_row, string i_column_name, string i_field_value, out string o_error)
        {
            o_error = "";

            int index_first_row = 0;
            Row first_row = GetRow(index_first_row, out o_error);
            if ("" != o_error) return false;

            int col_index = GetColumnIndex(first_row, i_column_name);
            if (col_index < 0)
            {
                o_error = "Table.SetFieldString There is no column: " + i_column_name;
                return false;
            }

            return SetFieldString(i_index_row, col_index, i_field_value, out o_error);
        }

        /// <summary>Returns field value as a string. Field is defined by a row index and a field index</summary>
        public string GetFieldString(int i_index_row, int i_index_field, out string o_error)
        {
            o_error = "";
            string ret_string = "";

            Row current_row = GetRow(i_index_row, out o_error);
            if ("" != o_error) return ret_string;

            Field current_field = current_row.GetField(i_index_field, out o_error);
            if ("" != o_error) return ret_string;

            ret_string = current_field.FieldValue;

            return ret_string;
        }

        /// <summary>Returns field value as a string. Field is defined by a row index and the column name (field value in the first row)</summary>
        public string GetFieldString(int i_index_row, string i_column_name, out string o_error)
        {
            o_error = "";
            string ret_string = "";

            Row current_row = GetRow(i_index_row, out o_error);
            if ("" != o_error) return ret_string;

            int index_first_row = 0;
            Row first_row = GetRow(index_first_row, out o_error);
            if ("" != o_error) return ret_string;

            int col_index = GetColumnIndex(first_row, i_column_name);
            if (col_index < 0)
            {
                o_error = "Table.GetFieldString There is no column: " + i_column_name;
                return ret_string;
            }

            ret_string = GetFieldString(i_index_row, col_index, out o_error);
            if (o_error != "")
            {
                ret_string = "";
                return ret_string;
            }

            return ret_string;
        }

        /// <summary>Get and set number of rows</summary>
        public int NumberRows
        {
            get { return m_number_rows; }

            set { m_number_rows = value; }
        }


        /// <summary>Get and set number of columns</summary>
        public int NumberColumns
        {
            get { return m_number_colums; }

            set { m_number_colums = value; }
        }

        /// <summary>Get column index for a given column name (field value in the first row)</summary>
        /// <returns>Negative value for failure</returns>
        static public int GetColumnIndex(Row i_row, string i_column_name)
        {
            int col_index = -1;

            for (int i_col = 0; i_col < i_row.NumberColumns; i_col++)
            {
                string error_message = "";
                Field current_field = i_row.GetField(i_col, out error_message);
                if (error_message != "")
                {
                    return col_index;
                }

                if (i_column_name == current_field.FieldValue)
                {
                    col_index = i_col;
                    break;
                }
            }

            return col_index;
        }

        /// <summary>Adds (appends) an empty row to the table</summary>
        public bool AddEmptyRow(out int o_row_index, out string o_error)
        {
            o_error = "";

            Row empty_row = new Row();

            for (int i_field = 0; i_field < NumberColumns; i_field++)
            {
                string field_value = "";

                Field current_field = new Field(field_value);

                empty_row.AddField(current_field);
            }

            o_row_index = NumberRows;

            if (!AddRow(empty_row, out o_error)) return false;

            return true;
        }

        /// <summary>Adds (appends) a row to the table</summary>
        public bool AddRow(Row i_row, out string o_error)
        {
            o_error = "";

            if (m_number_colums > 0)
            {
                if (m_number_colums != i_row.NumberColumns)
                {
                    o_error = "Table.AddRow Number of fields is " + m_number_colums.ToString() + " in table but number of row fields is " + i_row.NumberColumns.ToString();
                    return false;
                }
            }

            m_array_rows.Add(i_row);

            m_rows = (Row[])m_array_rows.ToArray(typeof(Row));

            NumberRows = m_rows.Length;

            return true;
        }

        /// <summary>Get row for a given index</summary>
        public Row GetRow(int i_index_row, out string o_error)
        {
            Row ret_row = new Row();
            o_error = "";

            if (i_index_row >= 0 && i_index_row < m_rows.Length)
            {
                ret_row = m_rows[i_index_row];
            }
            else
            {
                o_error = "Index " + i_index_row.ToString() + " is not between 0 and " + (m_rows.Length - 1).ToString();
                return ret_row;
            }

            return ret_row;
        }


        /// <summary>Remove row</summary>
        public bool RemoveRow(int i_index_row, out string o_error)
        {
            o_error = "";

            if (0 == i_index_row)
            {
                o_error = "Table.RemoveRow Not allowed to delete the first row";
                return false;
            }

            try
            {
                m_array_rows.RemoveAt(i_index_row);
            }
            catch
            {
                o_error = "Table.RemoveRow Row index " + i_index_row + " is not between 0 and " + m_array_rows.Count.ToString();

                return false;
            }

            m_rows = (Row[])m_array_rows.ToArray(typeof(Row));
            
            m_number_rows = m_rows.Length;

            return true;
        }


        /// <summary>Inserts a row in the table</summary>
        public bool InsertRow(Row i_row, int i_index_insert, out string o_error)
        {
            o_error = "";

            if (m_number_rows == 0)
            {
                o_error = "Table.InsertRow Number of rows is 0";
                return false;
            }

            if (0 == i_index_insert)
            {
                o_error = "Table.InsertRow Not allowed to insert row before the first row";
                return false;
            }

            if (m_number_colums != i_row.NumberColumns)
            {
                o_error = "Table.InsertRow Number of fields is " + m_number_colums.ToString() + " in table but number of row fields is " + i_row.NumberColumns.ToString();
                return false;
            }

            try
            {
                m_array_rows.Insert(i_index_insert, i_row);
            }
            catch
            {
                o_error = "Table.InsertRow Row index " + i_index_insert + " is not between 0 and " + m_array_rows.Count.ToString();

                return false;
            }

            m_rows = (Row[])m_array_rows.ToArray(typeof(Row));

            m_number_rows = m_rows.Length;

            return true;
        }

        /// <summary>Move row</summary>
        public bool MoveRow(int i_index_row, int i_index_move_to, out string o_error)
        {
            o_error = "";

            if (0 == i_index_row)
            {
                o_error = "Table.MoveRow Not allowed to move the first row";
                return false;
            }

            if (0 == i_index_move_to)
            {
                o_error = "Table.MoveRow Not allowed to move row before the first row";
                return false;
            }

            Row row_to_move = GetRow(i_index_row, out o_error);
            if (o_error != "") return false;

            if (!RemoveRow(i_index_row, out o_error)) return false;

            if (!InsertRow(row_to_move, i_index_move_to, out o_error)) return false;

            return true;
        }

    } // Table

    /// <summary>Holds the data for a row in a table. Has functions to add, remove and retrieve data</summary>
    public class Row
    {
        /// <summary>List of fields</summary>
        private ArrayList m_array_fields = new ArrayList();

        /// <summary>All fields in the row</summary>
        private Field[] m_fields;

        private int m_number_colums = 0;

        /// <summary>Get number of columns</summary>
        public int NumberColumns
        {
            get { return m_number_colums; }
        }

        /// <summary>Constructor ...</summary>
        public Row()
        {

        }

        /// <summary>Adds a field to the row</summary>
        public void AddField(Field i_field)
        {
            m_array_fields.Add(i_field);

            m_fields = (Field[])m_array_fields.ToArray(typeof(Field));

            m_number_colums = m_fields.Length;
        }

        /// <summary>Remove field with a given index</summary>
        public bool RemoveField(int i_index_field, out string o_error)
        {
            o_error = "";

            try
            {
                m_array_fields.RemoveAt(i_index_field);
            }
            catch 
            {
                o_error = "Row.RemoveField Field index " + i_index_field + " is not between 0 and " + m_array_fields.Count.ToString();

                return false;
            }

            m_fields = (Field[])m_array_fields.ToArray(typeof(Field));

            m_number_colums = m_fields.Length;

            return true;
        }

        /// <summary>Insert a field where the position is defined by index</summary>
        public bool InsertField(int i_index_field, Field i_field, out string o_error)
        {
            o_error = "";

            try
            {
                m_array_fields.Insert(i_index_field, i_field);
            }
            catch
            {
                o_error = "Row.InsertField Field index " + i_index_field + " is not between 0 and " + m_array_fields.Count.ToString();

                return false;
            }

            m_fields = (Field[])m_array_fields.ToArray(typeof(Field));

            m_number_colums = m_fields.Length;

            return true;
        }

        /// <summary>Set field value for field defined by the index</summary>
        public bool SetFieldValue(int i_index_field, string i_value, out string o_error)
        {
            o_error = "";

            if (i_index_field >= 0 && i_index_field < m_fields.Length)
            {
                Field current_field = m_fields[i_index_field];

                current_field.FieldValue = i_value;
            }
            else
            {
                o_error = "Field index " + i_index_field.ToString() + " is not between 0 and " + (m_fields.Length - 1).ToString();
            }

            return true;
        }

        /// <summary>Get field for a given index</summary>
        public Field GetField(int i_index_field, out string o_error)
        {
            o_error = "";
            Field ret_field = new Field("FieldIsNotDefined");

            if (i_index_field >= 0 && i_index_field < m_fields.Length)
            {
                ret_field = m_fields[i_index_field];
            }
            else
            {
                o_error = "Field index " + i_index_field.ToString() + " is not between 0 and " + (m_fields.Length - 1).ToString();
                return ret_field;
            }

            return ret_field;
        }
    }

    /// <summary>Holds the data for a column in a table. Has functions to add and retrieve column data</summary>
    public class Column
    {
        /// <summary>List of fields</summary>
        private ArrayList m_array_fields = new ArrayList();

        /// <summary>All fields in the column</summary>
        private Field[] m_fields;

        /// <summary>Number of fields in the column</summary>
        private int m_number_fields = 0;

        /// <summary>Get number of fields in the columne</summary>
        public int NumberFields
        {
            get { return m_number_fields; }
        }

        /// <summary>Constructor ...</summary>
        public Column()
        {

        }

        /// <summary>Clears the column, i.e. deletes all fields</summary>
        public void Clear()
        {
            m_array_fields.Clear();

            m_fields = (Field[])m_array_fields.ToArray(typeof(Field));

            m_number_fields = m_fields.Length;
        }

        /// <summary>Appends a field to the column</summary>
        public void AppendField(Field i_field)
        {
            m_array_fields.Add(i_field);

            m_fields = (Field[])m_array_fields.ToArray(typeof(Field));

            m_number_fields = m_fields.Length;
        }


        /// <summary>Get field for a given index</summary>
        public Field GetField(int i_index_field, out string o_error)
        {
            o_error = "";
            Field ret_field = new Field("FieldIsNotDefined");

            if (i_index_field >= 0 && i_index_field < m_fields.Length)
            {
                ret_field = m_fields[i_index_field];
            }
            else
            {
                o_error = "Field index " + i_index_field.ToString() + " is not between 0 and " + (m_fields.Length - 1).ToString();
                return ret_field;
            }

            return ret_field;
        }

       
    }

    /// <summary>Holds header data for a field (a column). Has functions to set and retrieve data</summary>
    public class FieldHeader
    {
        /// <summary>Name of the field (column)</summary>
        private string m_field_name = "";

        /// <summary>Caption for the field (column)</summary>
        private string m_field_caption = "";

        /// <summary>Type of field (column)</summary>
        private FieldType m_field_type = FieldType.UNDEFINED;

        /// <summary>Help (tool tips) for the field (column)</summary>
        private string m_field_help = "";

        /// <summary>Constructor that sets the name of the field (column)</summary>
        public FieldHeader(string i_field_name)
        {
            this.m_field_name = i_field_name;
        }

        /// <summary>Field (column) name</summary>
        public string Name
        {
            get { return this.m_field_name; }

            set { this.m_field_name = value; }
        }

        /// <summary>Field (column) caption</summary>
        public string Caption
        {
            get { return this.m_field_caption; }

            set { this.m_field_caption = value; }
        }

        /// <summary>Field (column) type</summary>
        public FieldType Type
        {
            get { return this.m_field_type; }

            set { this.m_field_type = value; }
        }

        /// <summary>Field (column) help</summary>
        public string Help
        {
            get { return this.m_field_help; }

            set { this.m_field_help = value; }
        }
    }

    /// <summary>Holds the header data for the rows. Has functions to set and retrieve data</summary>
    public class RowHeader
    {
        /// <summary>List of fields</summary>
        private ArrayList m_array_header_fields = new ArrayList();

        /// <summary>All fields in the row</summary>
        private FieldHeader[] m_header_fields;

        private int m_number_colums = 0;

        /// <summary>Get number of columns</summary>
        public int NumberColumns
        {
            get { return m_number_colums; }
        }

        /// <summary>Constructor ...</summary>
        public RowHeader()
        {

        }

        /// <summary>Adds a field header to the header row</summary>
        public void AddFieldHeader(FieldHeader i_field)
        {
            m_array_header_fields.Add(i_field);

            m_header_fields = (FieldHeader[])m_array_header_fields.ToArray(typeof(FieldHeader));

            m_number_colums = m_header_fields.Length;
        }

        /// <summary>Get header field for a given index</summary>
        public FieldHeader GetFieldHeader(int i_index_field, out string o_error)
        {
            o_error = "";
            FieldHeader ret_header_field = null;

            if (i_index_field >= 0 && i_index_field < m_header_fields.Length)
            {
                ret_header_field = m_header_fields[i_index_field];
            }
            else
            {
                o_error = "Field index " + i_index_field.ToString() + " is not between 0 and " + (m_header_fields.Length - 1).ToString();
                return ret_header_field;
            }

            return ret_header_field;
        }

        /// <summary>Remove Field header</summary>
        public bool RemoveFieldHeader(int i_index_field, out string o_error)
        {
            o_error = "";

            try
            {
                m_array_header_fields.RemoveAt(i_index_field);
            }
            catch
            {
                o_error = "Table.RemoveRow Row index " + i_index_field + " is not between 0 and " + m_header_fields.Length.ToString();

                return false;
            }

            m_header_fields = (FieldHeader[])m_array_header_fields.ToArray(typeof(FieldHeader));

            m_number_colums = m_header_fields.Length;

            return true;
        }

        /// <summary>Insert field header</summary>
        public bool InsertFieldHeader(FieldHeader i_field_header, int i_index_insert, out string o_error)
        {
            o_error = "";

            try
            {
                m_array_header_fields.Insert(i_index_insert, i_field_header);
            }
            catch
            {
                o_error = "Table.InsertFieldHeader Field index " + i_index_insert + " is not between 0 and " + m_array_header_fields.Count.ToString();

                return false;
            }

            m_header_fields = (FieldHeader[])m_array_header_fields.ToArray(typeof(FieldHeader));

            m_number_colums = m_header_fields.Length;

            return true;
        }
    }

    /// <summary>Holds the data for a field in a row. Has functions to set and retrieve data</summary>
    public class Field
    {
        /// <summary>Holds the value of the field as a string</summary>
        private string m_field_value;

        /// <summary>Constructor that sets the field value</summary>
        public Field(string i_field_value)
        {
            m_field_value = i_field_value;
        }

        /// <summary>Get or set the field value</summary>
        public string FieldValue
        {
            get { return m_field_value; }

            set { m_field_value = value; }
        }
    }

    /// <summary>Field type enumerator</summary>
    public enum FieldType
    {
        /// <summary>Undefined</summary>
        UNDEFINED,
        /// <summary>String</summary>
        STRING,
        /// <summary>Integer</summary>
        INTEGER,
        /// <summary>Float</summary>
        FLOAT,
        /// <summary>Boolean</summary>
        BOOLEAN
    };


}
