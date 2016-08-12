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
            string currentSheet = "Configuration";
            Excel.Worksheet excelWorksheet =
                (Excel.Worksheet)excelSheets.get_Item(currentSheet);

            // access cell within sheet
            Excel.Range excelCell =
                  (Excel.Range)excelWorksheet.get_Range("A1", "A1");

            //var excelApp = new Excel.Application();
            //excelApp.Workbooks.Add();

            //excelApp.ActiveCell.FormulaR1C1 = "1";
            //excelApp.Range["A2"].Select();
            //excelApp.ActiveCell.FormulaR1C1 = "2";
            //excelApp.Range["A1:A2"].Select();
            //excelApp.Selection.AutoFill(
            //    Destination: excelApp.Range["A1:A10"],
            //    Type: Excel.XlAutoFillType.xlFillDefault);
            //excelApp.Range["A1:A10"].Select();

            //excelApp.Selection.Interior.Pattern = Excel.Constants.xlSolid;
            //excelApp.Selection.Interior.PatternColorIndex = Excel.Constants.xlAutomatic;
            //excelApp.Selection.Interior.ThemeColor = Excel.XlThemeColor.xlThemeColorAccent1;
            //excelApp.Selection.Interior.TintAndShade = 0.399945066682943;
            //excelApp.Selection.Interior.PatternTintAndShade = 0;

            //excelApp.Range["A1:A10"].Select();
            //excelApp.ActiveSheet.Shapes.AddChart.Select();
            //excelApp.ActiveChart.ChartType = Excel.XlChartType.xlConeColStacked;
            //excelApp.ActiveChart.SetSourceData(Source: excelApp.Range["Sheet1!$A$1:$A$10"]);

            //excelApp.Visible = true;

            //long rMin = 0;
            //long rMax = 0;
            //Range rngClear = default(Range);




            //long i = 0;
            //var oXL = new Microsoft.Office.Interop.Excel.Application();
            //Excel.Worksheet oWS = null;
            //long rMin = 0;
            //long rMax = 0;
            //Range rngClear = default(Range);

            //try
            //{
            //    oWS = oXL.Worksheets["Renewals"]; //oXL.Parent.Worksheets.Item["myXlSheet"];
            //    oWS.Activate();

            //    //var _with1 = oXL.Worksheets["Renewals"];
            //    rMin = 5;
            //    //First row of IP renewal data

            //    rMax = oWS.UsedRange.Rows.Count;
            //    for (i = rMin; i <= rMax; i++)
            //    {
            //        oWS.Rows.Delete(i);
            //    }

            //    //rMax = oWS.Range("A" + oWS.Rows.Count).End(-4162).Row;
            //    //Last row of IP renewal data
            //    //if (rMax >= 5)
            //    //    oWS.Range(oWS.Rows(rMin), oWS.Rows(rMax)).Delete();

            //}

            //catch (Exception ex)
            //{
            //    Console.Write(ex.Message);

            //}

            //finally
            //{
            //    //oWS.Close();
            //    oXL.Quit();
            //    ResetCurrentCulture();
            //}

        }
    }
}
