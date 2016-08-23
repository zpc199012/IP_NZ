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


namespace IP_NZ
{
    public partial class TestClearList : Form
    {
        private Excel.Application _app;
        private Excel.Workbooks _books;
        private Excel.Workbook _book;
        protected Excel.Sheets _sheets;
        protected Excel.Worksheet _sheet;

        public TestClearList()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            long i;
            long rMin = 5;
            long rMax = 38;
            OpenExcelWorkbook(@modGlobalvars.G_ExcelPath);
            _sheet = (Excel.Worksheet)_sheets[1];
            _sheet.Select(Type.Missing);
            //Excel.Range range = _sheet.get_Range("A1:A1", Type.Missing);

            for (i = rMin; i <= rMax; i++)
            {
                Excel.Range range = _sheet.get_Range("A5", Type.Missing).EntireRow;
                range.Delete(Excel.XlDeleteShiftDirection.xlShiftUp);
                NAR(range);

            }

            NAR(_sheet);
            CloseExcelWorkbook();
            NAR(_book);
            _app.Quit();
            NAR(_app);

        }

        protected void OpenExcelWorkbook(string fileName)
        {
            _app = new Excel.Application();
            _app.Visible = true;
            if (_book == null)
            {
                _books = _app.Workbooks;
                _book = _books.Open(fileName, Type.Missing, false, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                _sheets = _book.Worksheets;
            }
        }
        protected void CloseExcelWorkbook()
        {
            _book.Save();
            _book.Close(false, Type.Missing, Type.Missing);
        }
        protected void NAR(object o)
        {
            try
            {
                if (o != null)
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(o);
            }
            finally
            {
                o = null;
            }
        }


    }
}
