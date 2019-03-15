using System;
using System.ComponentModel;
using System.Data;
using System.Data.OleDb;
using System.Runtime.InteropServices;
using System.Threading;
using System.Windows.Forms;
using Exl = Microsoft.Office.Interop.Excel;

namespace poa
{
    public partial class Form1 : Form
    {
        public Form1() => InitializeComponent();

        private void Button1_Click(object sender, EventArgs e)
        {
            
            Exl.Application xlApp;
            Exl.Workbook xlWorkBook;
            Exl.Worksheet xlWorkSheet;
            Exl.Range range;

            string str;
            int rowCount;
            int columnCount;
            int row = 0;
            int column = 0;

            xlApp = new Exl.Application();
            //xlWorkBook = xlApp.Workbooks.Open(@"d:\csharp-Excel.xls", 0, true, 5, "", "", true, Exl.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
            xlWorkBook = xlApp.Workbooks.Open(@"C:\Udvikling\po\data\tst1.xlsx");
            xlWorkSheet = (Exl.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;
            row = range.Rows.Count;
            column = range.Columns.Count;

            //for (rowCount = 1; rowCount <= row; rowCount++)
            //{
            //    for (columnCount = 1; columnCount <= column; columnCount++)
            //    {
            //        str = (string)(range.Cells[rowCount, columnCount] as Exl.Range).Value2;
            //        MessageBox.Show(str);
            //    }
            //}

            //iterate over the rows and columns and print to the console as it appears in the file
            //excel is not zero based!!
            for (int i = 1; i <= row; i++)
            {
                for (int j = 1; j <= column; j++)
                {
                    //new line
                    if (j == 1)
                        Console.Write("\r\n");

                    //write the value to the console
                    if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                        Console.Write(range.Cells[i, j].Value2.ToString() + "\t");

                    //add useful things here!   
                    if (j == 1)
                        readerBox.Text += "\r\n";

                    //write the value to the textbox
                    if (range.Cells[i, j] != null && range.Cells[i, j].Value2 != null)
                        readerBox.Text += range.Cells[i, j].Value2.ToString() + "\t";


                    //progress
                    // Wait 100 milliseconds.
                    Thread.Sleep(100);
                    // Report progress.
                }
            }

            Console.Write("\r\n");
            Console.WriteLine("DONE READING FILE");

            //cleanup
            GC.Collect();
            GC.WaitForPendingFinalizers();

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);


            //string name = "Items";
            //string constr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" +
            //                    "C:\\Sample.xlsx" +
            //                    ";Extended Properties='Excel 12.0 XML;HDR=YES;';";

            //OleDbConnection con = new OleDbConnection(constr);
            //OleDbCommand oconn = new OleDbCommand("Select * From [" + name + "$]", con);
            //con.Open();

            //OleDbDataAdapter sda = new OleDbDataAdapter(oconn);
            //DataTable data = new DataTable();
            //sda.Fill(data);
            //dataGridView1.DataSource = data;

        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            
        }
        


    }
}
