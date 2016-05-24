using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.OleDb;
using System.IO;
using Word = Microsoft.Office.Interop.Word;


namespace Retrieve_data_from_excel
{ 
    
    public partial class Form1 : Form
    {
  
        private controller control;
        public Form1()
        {
            InitializeComponent();
        }
        
        private void Form1_Load(object sender, EventArgs e)
        {
            
           control = controller.ConstructObject();
            dataGridView1.AllowUserToOrderColumns = true;
        }
        public void getsheets()
        {
            comboBox2.Items.Clear();
            string stringconn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + textselect.Text + ";Extended Properties=\"Excel 8.0;HDR=Yes;\";";
            OleDbConnection conn = new OleDbConnection(stringconn);
            try {
                conn.Open();
                DataTable Sheets = conn.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, null);
                string worksheets = Sheets.Rows[0]["TABLE_NAME"].ToString();
                string sqlQuery = String.Format("SELECT * FROM [{0}]", worksheets);
                OleDbDataAdapter da = new OleDbDataAdapter(sqlQuery, conn);
                DataTable dt = new DataTable();
                da.Fill(dt);
                foreach (DataRow dr in Sheets.Rows)
                {
                    string sht = dr[2].ToString().Replace("'", "");
                    sht = sht.Substring(0, sht.Length - 1);
                    comboBox2.Items.Add(sht);
                }
            } catch(Exception ex)
            {
                MessageBox.Show("please exit the excel file to be opened by the program or select a excel file  ");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            OpenFileDialog opfd = new OpenFileDialog();
            if (opfd.ShowDialog() == DialogResult.OK)
                textselect.Text = opfd.FileName;
            getsheets();
        }

        private void updateview()
        {
            dataGridView1.DataSource = control.returnsheetTable();
            comboBox1.Items.Clear();
            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                string columnName = this.dataGridView1.Columns[i].Name;
                comboBox1.Items.Add(columnName);
            }
        }



        private void button2_Click(object sender, EventArgs e)
        {
            try
            {
                if (control.SetSheet(textselect.Text, comboBox2.Text))
                {
                    dataGridView1.DataSource = control.returnsheetTable();
                    comboBox1.Items.Clear();
                    for (int i = 0; i < dataGridView1.Columns.Count; i++)
                    {
                        string columnName = this.dataGridView1.Columns[i].Name;
                        comboBox1.Items.Add(columnName);
                    }
                }
                else {
                    MessageBox.Show("failed to read excel file or ");
                }

            }
            catch (Exception ex) { MessageBox.Show("unexpected error please check on the sheet selected "); }
   
        }

        private void Insert_Table_Word(int speakers)
        {
                Word._Application objApp;
                Word._Document objDoc;
            try
            {
                object objMiss = System.Reflection.Missing.Value;
                object objEndOfDocFlag = "\\endofdoc"; /* \endofdoc is a predefined bookmark */

                object objStartOfDocFlag = "\\startofdoc";

                OpenFileDialog opfd = new OpenFileDialog();
                if (opfd.ShowDialog() == DialogResult.OK)

                    control.SetWordPath(opfd.FileName);

                //Start Word and create a new document.
                objApp = new Word.Application();
                objApp.Visible = true;
                objDoc = objApp.Documents.Open(opfd.FileName, ref objMiss, ref objMiss, ref objMiss);


                Word.Table objTab1; //create table object
                Word.Range objWordRng = objDoc.Bookmarks.get_Item(ref objEndOfDocFlag).Range; //go to end of document
                objTab1 = objDoc.Tables.Add(objWordRng, speakers + 1, dataGridView1.Columns.Count, ref objMiss, ref objMiss);
                objTab1.Borders.Enable = 1;
                objTab1.AllowAutoFit = true;
                int iRow, iCols;
                string strText;

                for (iCols = 0; iCols < dataGridView1.Columns.Count; iCols++)
                {
                    string columnName = this.dataGridView1.Columns[iCols].Name;
                    objTab1.Cell(1, iCols + 1).Range.Text = columnName; //add some text to cell

                }
                objTab1.Rows[1].Range.Font.Bold = 1; //make first row of table BOLD

                int flag = 0;
                for (iRow = 0; iRow < speakers; iRow++)
                    for (iCols = 0; iCols < dataGridView1.Columns.Count; iCols++)
                    {
                        if (iRow == speakers)
                        {
                            break;
                        }

                        strText = dataGridView1.Rows[iRow].Cells[iCols].Value.ToString();

                        iRow = iRow + 1;
                        objTab1.Cell(iRow + 1, iCols + 1).Range.Text = strText; //add some text to cell
                        iRow = iRow - 1;
                        flag++;
                        if (flag == dataGridView1.Columns.Count)
                        {
                            iRow++;
                            iCols = -1;
                            flag = 0;
                        }
                    }


                objWordRng = objDoc.Bookmarks.get_Item(ref objEndOfDocFlag).Range;
                objWordRng.InsertParagraphAfter(); //put enter in document
                objWordRng.InsertAfter("                                                ");

                object szPath = opfd.FileName;
                objDoc.SaveAs(ref szPath);

                objApp.Quit();


            }
            catch (Exception ex)
            {
                 MessageBox.Show("Error occurred while creating table please check on the file and close it if it is opened ");
            }
            finally
            {
                //you can dispose object here
            }

            }
           
        

        private void button3_Click(object sender, EventArgs e)
        {
            double transition = 0;
            if (textBox4.Text == "") { }
            else
                try
                {
                    transition = Convert.ToInt32(textBox4.Text);

                }
                catch (Exception ex)
                {
                    MessageBox.Show("the value inserted for transition time isn't a number ");
                    return;
                }

            if (transition < 0)
            {
                MessageBox.Show("the value inserted for transition time is negative please insert it by positive real number , zero or leave it empty \n it will be initially equal to 0 ");
                return;
            }


            int speakers = 0;
            try
            {
                speakers = Convert.ToInt32(textBox3.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show("number of the presenter isn't inserted correctly please try again");
                return;
            }
            if (dataGridView1.Rows.Count < speakers)
            {
                MessageBox.Show("number of the presenter requested is larger than the number of data which can be inserted please try again");
                return;
            }
            if (speakers < 0)
            {
                MessageBox.Show("number of the presenter requested isn't realistic please try again ");
                return;
            }
            bool time = false;
            DateTime date = dateTimePicker1.Value;
            DateTime date_start;
            DateTime date_end;
            double duration = 0; ;
            if (textDuration.Text != "")
                try
                {
                    duration = Convert.ToDouble(textDuration.Text);
                    time = true;
                }
                catch (Exception ex) { MessageBox.Show("The value entered in the field is wrong please try again"); return; }
            if (duration < 0)
            {
                MessageBox.Show("The duration of the speech isn't positive real number");
                return;
            }

            List<string> timetable = new List<string>();
            if (time == true)
            {

                for (int j = 0; j < dataGridView1.Columns.Count; j++)
                {         
                        string cur_col = this.dataGridView1.Columns[j].Name;
                        if (cur_col == "Time")
                        {
                            MessageBox.Show("Column Time Already Exists,Please Check On It Or Rename The Column Time");
                            return;
                        }
                    
                }

                     dataGridView1.Columns.Add("Time", "Time");
                    timetable.Add("Time");
                for (int i = 0; i < speakers; i++)
                {
                    //My addition
                    int index = dataGridView1.Columns["Time"].Index;
                    //////////////////////////////////////////////////

                    timetable.Add(date.ToString("HH:mm") + " - " + date.AddMinutes(duration).ToString("HH:mm"));
                    date_start = date;
                    date_end = date.AddMinutes(duration);
                    date = date.AddMinutes(duration + transition);

                    //My addition
                    try
                    {
                        dataGridView1.Rows[i].Cells[index].Value = date_start.ToString("HH:mm") + " - " + date_end.ToString("HH:mm");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error occurred while executing code : " + ex.Message);
                    }
                    ////////////////////////////////////////////////
                }
            }
            Insert_Table_Word(speakers);
            if(time == true)
            dataGridView1.Columns.Remove("Time");

        }
        private void button4_Click(object sender, EventArgs e)
        {
            if (comboBox1.Text == "" || comboBox3.Text == "" )
                MessageBox.Show("please select the column and the setting for filtering ");
            else
            {
                try
                {
                    string choice = "";
                    if (comboBox3.Text == "Keep only")
                        choice = " = ";
                    else if (comboBox3.Text == "Discard only") choice = " <> ";
                    else
                    {
                        MessageBox.Show("error choice of operation isn't selected correctly ");
                        return;
                    }
                    control.setsheetdata((DataTable)(dataGridView1.DataSource));
                    if (control.Query(choice, comboBox1.Text, textBox1.Text))
                    {
                        updateview();
                    }
                    else MessageBox.Show("Filter failed please try again and check on the input");
                }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Error please check on your input");

                    }
            }
            
        }

        private void button5_Click(object sender, EventArgs e)
        {
            try {
               
                control.setsheetdata((DataTable)(dataGridView1.DataSource));
                String val = comboBox1.Text;
                if (control.RemovesheetCol(val))
                {
                    updateview();
                }
                else { MessageBox.Show("please check on your input"); }
            }
            catch(Exception ex)
            { MessageBox.Show("please check on your input"); }
        }


        private void button7_Click(object sender, EventArgs e)
        {
            control.UndoExcel();
            updateview();
        }

        private void button8_Click(object sender, EventArgs e)
        {
            if (control.Renamecol(comboBox1.Text, textBox5.Text))
                updateview();
            else MessageBox.Show("Error the rename of the column has failed. Please check on column name or the value inserted \n Because it might be a duplicate or excel sheet doesn't exists");
        }

        private void button9_Click(object sender, EventArgs e)
        {
            string message = "";
            for (int j = 0; j < dataGridView1.Columns.Count; j++)
                message = message + j + "  "+ this.dataGridView1.Columns[j].Name + "-" + this.dataGridView1.Columns[j].DisplayIndex + "\n";
            MessageBox.Show(message);
        } 
    }

 }

