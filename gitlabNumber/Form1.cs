using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;
using SomeDatainLibrary;
namespace gitlabNumber
{
    public partial class Form1 : Form
    {
        string[] ExcelNameofColumn = ClasswithExcel.ExcelNameofColumn;
       
        public Form1()
        {
            InitializeComponent();
            
        }

        private void InserttoExcel(object sender, EventArgs e)
        {
            
            
            Cursor.Current = Cursors.WaitCursor;
            Excel.Application xlApp = new Excel.Application();
            Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(@"C:\Users\Wen-Liang\Desktop\GitlabNumberNew.xlsx");
            Excel._Worksheet xlWorksheet = xlWorkbook.Worksheets[1];
            Excel.Range oRng,oRngBefore;
            
            xlWorksheet.Tab.Color = Color.Blue;
          
            
            int[,] saNumber = new int[5,1];
            saNumber[0,0] = Int32.Parse(textBox1.Text);
            saNumber[1,0] = Int32.Parse(textBox2.Text);
            saNumber[2,0] = Int32.Parse(textBox3.Text);
            saNumber[3,0] = Int32.Parse(textBox4.Text);
            saNumber[4,0] = Int32.Parse(textBox5.Text);



            //string dateinfristrow = comboBox1.Text + "1";
            string dateinfristrow = "";
            string sumstart = "";
            string sumend = "";
            string totalcell = "";
            
            for (int i = 0; i < ExcelNameofColumn.Length; i++)
            {
                dateinfristrow = ExcelNameofColumn[i] + "1";
                if(xlWorksheet.get_Range(dateinfristrow,dateinfristrow).Value != null)
                {
                    dateinfristrow = ExcelNameofColumn[i + 1] + "1";
                    sumstart = ExcelNameofColumn[i + 1] + "2";
                    sumend = ExcelNameofColumn[i + 1] + "6";
                    totalcell = ExcelNameofColumn[i + 1] + "7";
                    continue;
                }

                string columnName = "The column name you should select is " + ExcelNameofColumn[i];
                MessageBox.Show(columnName,"Close",MessageBoxButtons.OK,MessageBoxIcon.Information);
                break;
            } 
            /*
            for(int i = 0; i < comboBox1.Items.Count; i++)
            {
                if (xlWorksheet.get_Range(dateinfristrow, dateinfristrow).Value != null)
                {
                    //MessageBox.Show("This column has already exist, please select next column", "Close", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    //return;
                    dateinfristrow = comboBox1.Items[i + 1].ToString() + "1";
                    sumstart = comboBox1.Items[i+1].ToString() + "2";
                    sumend = comboBox1.Items[i+1].ToString() + "6";
                    totalcell = comboBox1.Items[i + 1].ToString() + "7";
                    continue;
                }
                string columnName = "The column name you should select is " + comboBox1.Items[i].ToString();
                MessageBox.Show(columnName, "Close", MessageBoxButtons.OK, MessageBoxIcon.Information);
                break;
            }
            */
            
            xlWorksheet.get_Range(dateinfristrow, dateinfristrow).Value = dateTimePicker1.Value.ToShortDateString();
            
            oRngBefore = xlWorksheet.get_Range(sumstart, sumend);
            oRngBefore.Value = saNumber;
            oRng = xlWorksheet.get_Range(totalcell, totalcell);
            oRng.Formula = "=SUM(" + sumstart + ":" + sumend +")";
            xlApp.Visible = false;
            xlApp.UserControl = false;
            xlWorkbook.Save();

            xlWorkbook.Close();
            xlApp.Quit();

            Cursor.Current = Cursors.Default;

            
            
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var result = MessageBox.Show("Do you really want to close the gitlabNumber?", "Close", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                this.Close();
            }
            //this.Close();
        }
    }
}
