using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
namespace Retrieve_data_from_excel
{
    class controller
    {
        private ExcelSheet sheet;
        private WordFile word;
        private static int count = 0;
        private static controller cont;
        private controller() {
            sheet = ExcelSheet.ConstructObject();
            word = WordFile.ConstructObject();
        }
        ~controller() { count = 0; }
        public static controller ConstructObject()
        {
            if (count == 0)
            {
                cont = new controller();
                count++;
                return cont;
            }
            else
            { return cont; }
        }
        public bool SetSheet ( string select , string sheetnum )
            {
            return sheet.createconnection(select, sheetnum);
               
            } 
       public void SetWordPath(string name)
        {
            word.SetPath(name);
        }
        public string GetwordPath()
        {
            return word.GetPath();
                }
        public void setsheetdata(DataTable db)
        {
            sheet.Setdata(db);
        }
        public bool RemovesheetCol(string col)
        {
            return sheet.RemoveCol(col);
        }
        public DataTable returnsheetTable()
        { return sheet.GetDatatable(); }
        public bool Query(string choice, string col, string val)
        {
            return sheet.Query(choice, col, val);
        }
        public void UndoExcel()
        {
            sheet.Undo();
        }
        public bool Renamecol(string col , string rename)
        {
            return sheet.Renamedcolumn( col,  rename);
        }

    }
}
