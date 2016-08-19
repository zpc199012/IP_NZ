using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.VisualBasic;
using System.Collections;
using System.Diagnostics;
using System.IO;

using Microsoft.Office.Interop;

using System.Threading;
using System.Text.RegularExpressions;
using System.Runtime.InteropServices;

namespace IP_NZ
{
    public partial class frmMain : Form
    {
        public frmMain()
        {
            InitializeComponent();
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            string dateTimePicker;
            System.Globalization.CultureInfo ci = new System.Globalization.CultureInfo("en-US", false);
            System.Globalization.CultureInfo newCi = (System.Globalization.CultureInfo)ci.Clone();
            newCi.DateTimeFormat.AMDesignator = "AM";
            newCi.DateTimeFormat.PMDesignator = "PM";
            newCi.DateTimeFormat.ShortDatePattern = "MM/dd/yyyy";
            Thread.CurrentThread.CurrentCulture = newCi;
            modGlobalvars.G_LOADate = GetMaxLOADate();

            dateTimePicker = modGlobalvars.G_LOADate.Substring(4, 2) + "/" + modGlobalvars.G_LOADate.Substring(6, 2) + "/" + modGlobalvars.G_LOADate.Substring(0, 4);

            this.txtLOADate.Value = DateTime.Parse(dateTimePicker);

            this.Text = "IP Renewal Utility for " + modGlobalvars.G_ctry;
        }

        private void lblLOADate_Click(object sender, EventArgs e)
        {

        }

        private void Startbutton_Click(object sender, EventArgs e)
        {

            Process objProc = new Process();
            long retCtr = 0;
            long recordCount = 0;
            string yyyymmdd = string.Empty;

            yyyymmdd = Strings.Format(this.txtLOADate.Value, "yyyyMMdd");

            recordCount = GetRecordCount(yyyymmdd);
            if (recordCount > 0)
            {
                retCtr = PopulateExcel(modGlobalvars.G_ExcelPath, yyyymmdd, recordCount);

                Interaction.MsgBox("Number of cases added = " + retCtr, MsgBoxStyle.Information, modGlobalvars.G_ctry + " IP Renewal Message");
                this.ProgressBar1.Visible = false;
                this.lblprogress1.Visible = false;
                objProc.StartInfo.FileName = modGlobalvars.G_ExcelPath;
                objProc.StartInfo.Arguments = "";

                this.Visible = false;
                this.ShowInTaskbar = false;



                objProc.Start();

                //Test Form Start
                //TestClearList ClearList = new TestClearList();
                //ClearList.Show();

                frmService Service = new frmService();
                Service.Show();
                //Test End

                //objProc.Close();

                //while (!(objProc.HasExited))
                //{

                //}
                //Interaction.MsgBox("Process Done.", MsgBoxStyle.Information, modGlobalvars.G_ctry + " IP Renewal Message");
            }
            else
            {
                Interaction.MsgBox("No cases added with LOA Date=" + yyyymmdd, MsgBoxStyle.Information, modGlobalvars.G_ctry + " IP Renewal Message");
            }
            this.Visible = true;
            this.ShowInTaskbar = true;
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            modGlobalvars.G_cnn.Close();
            Application.Exit();
        }

        private void frmMain_FormClosed(object sender, System.Windows.Forms.FormClosedEventArgs e)
        {
            modGlobalvars.G_cnn.Close();
            Application.Exit();
        }

        private long PopulateExcel(string FullFileName, string yyyymmdd, long totalrows)
        {
            System.Data.Odbc.OdbcCommand objQuery = new System.Data.Odbc.OdbcCommand("select patno,rentyp,loadat,rensts,postdat from " + modGlobalvars.G_library.Trim() + ".APFRENEW WHERE CTRY = '" + modGlobalvars.G_ctry + "' AND LOADAT = '" + Strings.Trim(yyyymmdd) + "'", modGlobalvars.G_cnn);
            System.Data.Odbc.OdbcDataReader odbcReader = objQuery.ExecuteReader();

            //Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
            //Microsoft.Office.Interop.Excel.Workbook oWBK = null;
            //Microsoft.Office.Interop.Excel._Worksheet oWS = null;

            //Modified
            //Microsoft.Office.Interop.Range 

            IntPtr processId = default(IntPtr);
            string ss = string.Empty;
            //            long lngCtr = 0;
            //            long lngMaxCtr = 0;
            long retRowCtr = 0;
            int progressCtr = 0;
            string strRENTYP = string.Empty;
            string pattern = "(NZ?)";
            // Get our Excel process ID
            GetWindowThreadProcessId(modGlobalvars.oXL.Hwnd, ref processId);
            Process excelProcess = Process.GetProcessById(processId.ToInt32());

            try
            {
                modGlobalvars.oWBK = modGlobalvars.oXL.Workbooks.Open(FullFileName);
                modGlobalvars.oWS = modGlobalvars.oXL.Worksheets["Renewals"];
                modGlobalvars.oWS.Activate();
                if (!string.IsNullOrEmpty(Strings.Trim(modGlobalvars.oXL.Cells[5, 1].value)))
                {
                    modGlobalvars.oXL.Run("ClearIPList");
                }
                dynamic intXLSDataCtr = 0;
                int rowIndex = 1;
                //                int colIndex = 1;
                int rowLimit = modGlobalvars.oWS.UsedRange.Rows.Count;
                int colLimit = modGlobalvars.oWS.UsedRange.Columns.Count;

                rowLimit = modGlobalvars.oXL.ActiveSheet.Cells(modGlobalvars.oXL.Rows.Count, "A").End(-4162).Row;

                rowIndex = rowLimit + 1;
                progressCtr = 0;
                this.ProgressBar1.Minimum = 0;
                this.ProgressBar1.Maximum = (int)totalrows;
                this.ProgressBar1.Value = 0;
                this.ProgressBar1.Visible = true;
                this.lblprogress1.Visible = true;

                while (odbcReader.Read())
                {
                    //Me.ProgressBar1.Value += 1
                    progressCtr += 1;
                    this.ProgressBar1.PerformStep();

                    this.lblprogress1.Text = Conversion.Str(Math.Round((progressCtr / totalrows) * 100.0)) + "% Completed";
                    this.lblprogress1.Refresh();

                    if (((Microsoft.Office.Interop.Excel.Range)modGlobalvars.oXL.Cells[rowIndex, 1]).Value2 == null)
                    {
                        switch (odbcReader.GetString(1))
                        {
                            case "DES":
                                strRENTYP = "D";
                                break;
                            case "PAT":
                                strRENTYP = "P";
                                break;
                            default:
                                strRENTYP = "T";
                                break;
                        }
                        modGlobalvars.oXL.Cells[rowIndex, 1].value = strRENTYP;
                    }

                    if (((Microsoft.Office.Interop.Excel.Range)modGlobalvars.oXL.Cells[rowIndex, 2]).Value2 == null)
                    {
                        //oXL.Cells(rowIndex, 2).value = odbcReader.GetString("0")
                        modGlobalvars.oXL.Cells[rowIndex, 2].value = Regex.Replace(odbcReader.GetString(0), pattern, string.Empty);
                    }

                    retRowCtr += 1;
                    rowIndex += 1;

                }
            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);

            }
            finally
            {
                modGlobalvars.oWBK.Save();
                modGlobalvars.oWBK.Close();
                modGlobalvars.oXL.Quit();
                objQuery.Dispose();
                odbcReader.Close();
                //Make sure that our Excel process is removed
                excelProcess.Kill();

                modGlobalvars.oXL = null;
                modGlobalvars.oWBK = null;
                modGlobalvars.oWS = null;
            }

            return retRowCtr;
        }

        private long GetRecordCount(string yyyymmdd)
        {
            System.Data.Odbc.OdbcCommand objQuery = new System.Data.Odbc.OdbcCommand("select count(*) from " + modGlobalvars.G_library.Trim() + ".APFRENEW WHERE CTRY = '" + modGlobalvars.G_ctry + "' And LOADAT = '" + Strings.Trim(yyyymmdd) + "'", modGlobalvars.G_cnn);
            System.Data.Odbc.OdbcDataReader odbcReader = objQuery.ExecuteReader();
            long totalrows = 0;

            while (odbcReader.Read())
            {
                totalrows = Convert.ToInt32(odbcReader.GetString(0));

            }
            odbcReader.Close();

            return totalrows;
        }

        private string GetMaxLOADate()
        {
            System.Data.Odbc.OdbcCommand objQuery = new System.Data.Odbc.OdbcCommand("select MAX(LOADAT) from " + modGlobalvars.G_library.Trim() + ".APFRENEW WHERE CTRY = '" + modGlobalvars.G_ctry + "' And RENSTS = ' '", modGlobalvars.G_cnn);
            System.Data.Odbc.OdbcDataReader odbcReader = objQuery.ExecuteReader();
            string maxLOADate = string.Empty;

            while (odbcReader.Read())
            {
                maxLOADate = odbcReader.GetString(0);

            }
            odbcReader.Close();

            return maxLOADate;

        }

        [DllImport("user32", CharSet = CharSet.Ansi, SetLastError = true, ExactSpelling = true)]

        private static extern IntPtr GetWindowThreadProcessId(int hWnd, ref IntPtr lpdwProcessId);

        private void txtLOADate_GotFocus(object sender, System.EventArgs e)
        {
            this.txtLOADate.CustomFormat = "MM/dd/yyyy";
        }

        //private void clearList_Click(object sender, EventArgs e)
        //{
        //    System.Data.Odbc.OdbcCommand objQuery = new System.Data.Odbc.OdbcCommand("select patno,rentyp,loadat,rensts,postdat from " + modGlobalvars.G_library.Trim() + ".APFRENEW WHERE CTRY = '" + modGlobalvars.G_ctry + "' AND LOADAT = '" + Strings.Trim(yyyymmdd) + "'", modGlobalvars.G_cnn);
        //    System.Data.Odbc.OdbcDataReader odbcReader = objQuery.ExecuteReader();

        //    long i = 0;
        //    //Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
        //    //Microsoft.Office.Interop.Excel._Worksheet oWS = null;
        //    long rMin = 0;
        //    long rMax = 0;
        //    //Range rngClear = default(Range);

        //    try
        //    {
        //        modGlobalvars.oWS = modGlobalvars.oXL.Worksheets["Renewals"];
        //        modGlobalvars.oWS.Activate();

        //        //var _with1 = oXL.Worksheets["Renewals"];
        //        rMin = 5;
        //        //First row of IP renewal data

        //        rMax = modGlobalvars.oWS.UsedRange.Rows.Count;
        //        for (i = rMin; i <= rMax; i++)
        //        {
        //            modGlobalvars.oWS.Rows.Delete(i);
        //        }

        //        //rMax = oWS.Range("A" + oWS.Rows.Count).End(-4162).Row;
        //        //Last row of IP renewal data
        //        //if (rMax >= 5)
        //        //    oWS.Range(oWS.Rows(rMin), oWS.Rows(rMax)).Delete();

        //    }

        //    catch (Exception ex)
        //    {
        //        Console.Write(ex.Message);

        //    }
        //    finally
        //    {
        //        modGlobalvars.oWBK.Save();
        //        modGlobalvars.oWBK.Close();
        //        modGlobalvars.oXL.Quit();
        //        objQuery.Dispose();
        //        odbcReader.Close();
        //        //Make sure that our Excel process is removed
        //        //excelProcess.Kill();

        //        modGlobalvars.oXL = null;
        //        modGlobalvars.oWBK = null;
        //        modGlobalvars.oWS = null;
        //    }

        //}




        public string yyyymmdd { get; set; }
    }
}
