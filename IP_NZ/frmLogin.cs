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
using Microsoft.VisualBasic.ApplicationServices;
using System.Collections;
using System.Diagnostics;

using System.IO;
using System.Net.Mail;
using Microsoft.Office.Interop;

using IP_NZ;

namespace IP_NZ
{


    public partial class frmLogin : Form
    {

        private void OK_Click(System.Object sender, System.EventArgs e)
        {
            if (UserNamePassword())
            {
                modGlobalvars.G_user = UsernameTextBox.Text;
                modGlobalvars.G_pwd = PasswordTextBox.Text;
                this.Visible = false;
                this.ShowInTaskbar = false;

                ////Test Form Start
                //frmService Service = new frmService();
                //Service.Show();
                ////Test End

                frmMain frm = new frmMain();
                frm.Show();
            }
        }

        private void Cancel_Click(System.Object sender, System.EventArgs e)
        {
            this.Close();
        }

        private bool UserNamePassword()
        {

            if (string.IsNullOrEmpty(UsernameTextBox.Text))
            {
                //Using System.Windows.Forms.MessageBox.Show() instead of Interaction.MsgBox
                Interaction.MsgBox("User must not be blank", MsgBoxStyle.Critical, "User Error");
                return false;
            }

            if (string.IsNullOrEmpty(PasswordTextBox.Text))
            {
                //Using System.Windows.Forms.MessageBox.Show() instead of Interaction.MsgBox
                Interaction.MsgBox("Password must not be blank", MsgBoxStyle.Critical, "Password Error");
                return false;
            }
            modGlobalvars.G_connection = "Driver={Client Access ODBC Driver (32-bit)};" + "System=" + modGlobalvars.G_system.Trim() + ";" + "TRANSLATE=1;" + "Uid=" + UsernameTextBox.Text.Trim() + ";" + "Pwd=" + PasswordTextBox.Text.Trim();

            if (Convert.ToInt32(connServer(modGlobalvars.G_connection)) < 1)
            {
                return false;
            }

            return true;
        }

        protected object connServer(string strConn)
        {
            modGlobalvars.G_cnn = new System.Data.Odbc.OdbcConnection(strConn);
            Cursor.Current = Cursors.WaitCursor;
            try
            {
                modGlobalvars.G_cnn.Open();
                // will generate an error

            }
            catch (Exception ex)
            {
                Cursor.Current = Cursors.Default;

                // For MessageBox.Show InStr function, try "System.Environment.NewLine"
                Interaction.MsgBox(Strings.Mid(ex.Message, 1, Strings.InStr(ex.Message, Constants.vbCrLf)), MsgBoxStyle.Critical, "Error Message");
                modGlobalvars.G_cnn.Close();
                modGlobalvars.G_cnn.Dispose();
                return false;
            }
            Cursor.Current = Cursors.Default;
            return true;

        }


        public frmLogin()
        {
            InitializeComponent();
        }

        private void frmLogin_Load(object sender, EventArgs e)
        {
            Microsoft.VisualBasic.ApplicationServices.WindowsFormsApplicationBase
            myApplication = new WindowsFormsApplicationBase();
            modGlobalvars.G_system = myApplication.CommandLineArgs[0];
            modGlobalvars.G_library = myApplication.CommandLineArgs[1];
            modGlobalvars.G_ExcelPath = myApplication.CommandLineArgs[2];
            modGlobalvars.G_ctry = myApplication.CommandLineArgs[3];

        }


    }

}
