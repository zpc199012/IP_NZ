using Microsoft.Office.Interop.Excel;
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
//using Microsoft.Office.Interop.Excel.Worksheet;


namespace IP_NZ
{
    public partial class frmService : Form
    {
        public frmService()
        {
            InitializeComponent();
            //objProc.Start();
            //frmMain.Visible = true;
            //frmMain.ShowInTaskbar = true;
        }

        //Modified Start 08092016
        System.Globalization.CultureInfo oldCI;

        void SetNewCurrentCulture()
        {
            oldCI = System.Threading.Thread.CurrentThread.CurrentCulture;
            System.Threading.Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US");
        }

        void ResetCurrentCulture()
        {
            System.Threading.Thread.CurrentThread.CurrentCulture = oldCI;
        }

        //Modified End

        private void RenewalLabel_Click(object sender, EventArgs e)
        {

        }

        //private void clearList_Click(object sender, EventArgs e)
        //{
        //    Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();

        //    var rng = oXL.Worksheets["Renewals"]; 
        //    rng.Clear();
        //}

        private void clearList_Click(object sender, EventArgs e)
        {
            long i = 0;
            //var oXL = new Microsoft.Office.Interop.Excel.Application();
            Excel.Worksheet oWS = null;
            long rMin = 0;
            long rMax = 0;
            Range rngClear = default(Range);


            Excel.Application excelApp = new Excel.Application();

            // if you want to make excel visible to user, set this property to true, false by default
            excelApp.Visible = true;

            // open an existing workbook
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Open(modGlobalvars.G_ExcelPath,
                0, false, 5, "", "", false, Excel.XlPlatform.xlWindows, "",
                true, false, 0, true, false, false);

            // get all sheets in workbook
            Excel.Sheets excelSheets = excelWorkbook.Worksheets;

            // get some sheet
            string currentSheet = "Renewals";
            Excel.Worksheet excelWorksheet =
                (Excel.Worksheet)excelSheets.get_Item(currentSheet);

            // access cell within sheet
            Excel.Range excelCell =
                  (Excel.Range)excelWorksheet.get_Range("A5", "B25");

            
            try
            {
                //oWS = oXL.Worksheets["Renewals"]; //oXL.Parent.Worksheets.Item["myXlSheet"];
                excelWorksheet.Activate();

                //var _with1 = oXL.Worksheets["Renewals"];
                rMin = 5;
                //First row of IP renewal data

                rMax = excelWorksheet.UsedRange.Rows.Count;

                for (i = rMin; i <= rMax; i++)
                {
                    //Excel.Range cel = excelWorksheet.Cells[6, 2];
                    //cel.Delete();

                    //excelWorksheet.Range("A5", "B10");
                    excelWorksheet.Rows.Delete(i);
                }

                //rMax = oWS.Range("A" + oWS.Rows.Count).End(-4162).Row;
                //Last row of IP renewal data
                //if (rMax >= 5)
                //    oWS.Range(oWS.Rows(rMin), oWS.Rows(rMax)).Delete();

            }

            catch (Exception ex)
            {
                Console.Write(ex.Message);

            }

            finally
            {
                //oWS.Close();
                excelApp.Quit();
                //ResetCurrentCulture();
            }

        }
    }
}
