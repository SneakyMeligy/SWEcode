using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;
using System.Data.OleDb;
using System.IO;
using System.Data;

namespace Retrieve_data_from_excel
{
    class ExcelSheet
    {
        private ExcelSheet() { }
        ~ExcelSheet() { ExcelSheetcount = 0; }
        private static int ExcelSheetcount = 0;
        private static ExcelSheet Excel;
        private string path;
        private DataTable dt;
        private DataTable intdt;
        private OleDbConnection conn;
        string sheet;
        List<string> choices = new List<string>();
        List<string> queryval = new List<string>();
        List<string> columnchoice = new List<string>();
        public static ExcelSheet ConstructObject()
        {
            if (ExcelSheetcount == 0)
            {
                Excel = new ExcelSheet();
                ExcelSheetcount++;
                return Excel;
            }
            else
            { return Excel; }
        }
        public OleDbConnection GetConn()
        { return conn; }
        public bool createconnection( string path , string sheet )
        {
            List < string > columnsnames = new List<string>();
            bool success = false;
            queryval.Clear();
            choices.Clear();
            columnchoice.Clear();
            try
            {
                string stringconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path + @";Extended Properties=""Excel 8.0;IMEX=1;HDR=NO;TypeGuessRows=0;ImportMixedTypes=Text""";
                conn = new OleDbConnection(stringconn);
                OleDbDataAdapter da = new OleDbDataAdapter("Select * from [" + sheet + "$]", conn);

                dt = new DataTable();
                da.Fill(dt);
                
               
                for (int i = 0; i < dt.Columns.Count; i++)
                {
                    string search = dt.Rows[0][i].ToString();
                    if (search == "")
                    {
                        search = " ";
                    }
                    while ((columnsnames.Exists(e => e.Equals(search))))
                        search = search + i.ToString();
                    
                    dt.Columns[i].ColumnName = search;
                    columnsnames.Add(dt.Columns[i].ColumnName);
                }
                dt.Rows[0].Delete();
                dt.AcceptChanges();
                for (int i = dt.Rows.Count - 1; i >= 0; i--)
                {
                    bool delete = true;
                    for(int j = dt.Columns.Count - 1; j >= 0; j--)
                    {
                        if (dt.Rows[i][j].ToString() != "")
                        {
                            delete = false;
                            break;
                        }
                    }
                    if(delete == true)
                        dt.Rows[i].Delete();
                }

                dt.AcceptChanges();

                foreach (var column in dt.Columns.Cast<DataColumn>().ToArray())
                {
                    if (dt.AsEnumerable().All(dr => dr.IsNull(column)))
                        dt.Columns.Remove(column);
                }
                intdt = dt.Copy();
                success = true;
            } catch(Exception)
            { success = false; }
            return success;
        }
        public bool RemoveCol(string col)
        {
           
            bool success = false;
            try {
                dt.Columns.Remove(col);
                success = true;
                choices.Add("col");
                queryval.Add(" ");
                columnchoice.Add(col);
            } catch(Exception)
            { success = false; }
            dt.AcceptChanges();
            return success;
        }

        public bool Renamedcolumn(string col, string rename)
        {
            if (rename == "")
                rename = " ";

            if (dt == null)
                return false;
                for (int j = 0; j < dt.Columns.Count; j++)
                {
                if (dt.Columns[j].ColumnName == rename)
                    return false;
                }
            for (int i = 0; i < dt.Columns.Count;  i++)
            {
                if (dt.Columns[i].ColumnName == col)
                {
                    dt.Columns[i].ColumnName = rename;
                    choices.Add("rename col");
                    queryval.Add(rename);
                    columnchoice.Add(col);
                    return true;
                }
            }

            return false;

        }

        public void Setdata(DataTable data)
        { dt = data; }
        public DataTable GetDatatable()
        { return dt;  }
        public string GetPath()
        { return path; }

        public bool Query(string choice , string col , string val )
        {
            bool success = false;
            try
            {

                if (choice == " <> ")
                {
                    for (int i = dt.Rows.Count - 1; i >= 0; i--)
                    {
                        // whatever your criteria is
                        if (dt.Rows[i][col].ToString() == val)
                            dt.Rows[i].Delete();
                    }
                    queryval.Add(val);
                    choices.Add(choice);
                    columnchoice.Add(col);
                    dt.AcceptChanges();
                    return true;
                }
                else if (choice == " = ")
                {
                    for (int i = dt.Rows.Count - 1; i >= 0; i--)
                    {
                        // whatever your criteria is
                        if (dt.Rows[i][col].ToString() != val)
                            dt.Rows[i].Delete();
                    }
                    queryval.Add(val);
                    choices.Add(choice);
                    columnchoice.Add(col);
                    dt.AcceptChanges();
                    return true;
                }

            }
            catch (Exception) { }
            return success;
        }


        public void Undo()
        {
            if (choices.Count == 0)
                return;
            choices.RemoveAt(choices.Count - 1);
            queryval.RemoveAt(queryval.Count - 1);
            columnchoice.RemoveAt(columnchoice.Count - 1);
            dt = intdt.Copy();
            for (int i = 0; i < choices.Count; i++)
            {
                if (choices[i] == "col")
                {
                    RemoveCol(columnchoice[i]);
                    choices.RemoveAt(choices.Count - 1);
                    queryval.RemoveAt(queryval.Count - 1);
                    columnchoice.RemoveAt(columnchoice.Count - 1);
                }
                else if (choices[i] == "rename col") {
                    Renamedcolumn(columnchoice[i], queryval[i]);
                    choices.RemoveAt(choices.Count - 1);
                    queryval.RemoveAt(queryval.Count - 1);
                    columnchoice.RemoveAt(columnchoice.Count - 1);
                }
                else 
                {
                    Query(choices[i], columnchoice[i], queryval[i]);
                    choices.RemoveAt(choices.Count - 1);
                    queryval.RemoveAt(queryval.Count - 1);
                    columnchoice.RemoveAt(columnchoice.Count - 1);
                }
            }
        }

    }
}

