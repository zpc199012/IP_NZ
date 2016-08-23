using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace IP_NZ
{
    class modGlobalvars
    {
        public static string G_connection = string.Empty;
        public static System.Data.Odbc.OdbcConnection G_cnn;
        public static string G_user = string.Empty;
        public static string G_pwd = string.Empty;
        public static string G_system = string.Empty;
        public static string G_library = string.Empty;
        public static string G_ExcelPath = string.Empty;
        public static string G_ctry = string.Empty;
        public static string G_LOADate = string.Empty;

        public static Microsoft.Office.Interop.Excel.Application oXL = new Microsoft.Office.Interop.Excel.Application();
        public static Microsoft.Office.Interop.Excel.Workbook oWBK = null;
        public static Microsoft.Office.Interop.Excel._Worksheet oWS = null;
    }
}
