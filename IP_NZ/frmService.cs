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
            Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel._Worksheet oWS = null;
            long rMin = 0;
            long rMax = 0;
            Range rngClear = default(Range);

            try
            {
                oWS = oXL.Worksheets["Renewals"];
                oWS.Activate();

                //var _with1 = oXL.Worksheets["Renewals"];
                rMin = 5;
                //First row of IP renewal data

            //var _with1 = oXL.Worksheets["Renewals"];
            rMin = 5;
            //First row of IP renewal data
                //Last row of IP renewal data
                //if (rMax >= 5)
                //    oWS.Range(oWS.Rows(rMin), oWS.Rows(rMax)).Delete();

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);

            }

        }
    }
}
