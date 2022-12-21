using System;
using System.IO;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using ClosedXML.Excel;
using Excel2 = Microsoft.Office.Interop.Excel;
using ExcelDataReader;



namespace datadisplay
{
    public partial class Form1 : Form
    {


        public Form1()
        {
            InitializeComponent();
        }

        private void readButton_Click(object sender, EventArgs e)
        {
            using (OpenFileDialog dialog = new OpenFileDialog()
            { Filter = "Excel workbook|*.xlsx", Multiselect = false })
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    //Define the datatable to store the data from excel
                    DataTable dt = new DataTable();

                    using (XLWorkbook workbook = new XLWorkbook(dialog.FileName))
                    {
                        bool IsFirstRow = true;     //Here we check we are in the process of writing the header
                        var rows = workbook.Worksheet(3).RowsUsed();    //Choosing sheet 3 from the workbook
                        foreach (var row in rows)
                        {
                            if (IsFirstRow)
                            {
                                //adding columns
                                foreach (IXLCell cell in row.Cells())
                                {
                                    dt.Columns.Add(cell.Value.ToString());
                                }
                                IsFirstRow = false;
                            }
                            else
                            {
                                dt.Rows.Add();
                                int i = 0;
                                foreach (IXLCell cell in row.Cells())
                                    dt.Rows[dt.Rows.Count - 1][i++] = cell.Value.ToString();
                            }
                        }
                        //Show data on data grid view
                        dataGrid.DataSource = dt;

                    }

                }
            }
        }

        private void buttonSave_Click(object sender, EventArgs e)
        {
            Excel2._Application app = new Excel2.Application();
            Excel2._Workbook workbook = app.Workbooks.Add(Type.Missing);
            Excel2._Worksheet worksheet = null;

            app.Visible = true;
            worksheet = workbook.Sheets["Sheet1"];
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "EditSheet1";
            

            for (int i = 1; i < dataGrid.Columns.Count + 1; i++)
            {
                worksheet.Cells[1,i] = dataGrid.Columns[i-1].HeaderText;
            }

            for (int i = 1; i < dataGrid.Rows.Count - 1; i++)
            {
                for (int j = 1; j < dataGrid.Columns.Count; j++)
                {
                    worksheet.Cells[i+1, j] = dataGrid.Rows[i-1].Cells[j-1].Value.ToString();
                }
            }
            workbook.SaveAs("C:\\Users\\Dominic Vuga\\Desktop\\myVSprojects\\datadisplay\\test.csv");
        }
    }
}
