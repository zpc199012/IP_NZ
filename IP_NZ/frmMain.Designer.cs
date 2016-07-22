namespace IP_NZ
{
    partial class frmMain
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.lblLOADate = new System.Windows.Forms.Label();
            this.txtLOADate = new System.Windows.Forms.DateTimePicker();
            this.Startbutton = new System.Windows.Forms.Button();
            this.CloseButton = new System.Windows.Forms.Button();
            this.ProgressBar1 = new System.Windows.Forms.ProgressBar();
            this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
            this.lblprogress1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // lblLOADate
            // 
            this.lblLOADate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.lblLOADate.ForeColor = System.Drawing.SystemColors.ControlText;
            this.lblLOADate.Location = new System.Drawing.Point(97, 78);
            this.lblLOADate.Name = "lblLOADate";
            this.lblLOADate.Size = new System.Drawing.Size(142, 21);
            this.lblLOADate.TabIndex = 35;
            this.lblLOADate.Text = "Enter LOA Date:";
            this.lblLOADate.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.lblLOADate.Click += new System.EventHandler(this.lblLOADate_Click);
            // 
            // txtLOADate
            // 
            this.txtLOADate.CalendarFont = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.txtLOADate.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.txtLOADate.Format = System.Windows.Forms.DateTimePickerFormat.Custom;
            this.txtLOADate.Location = new System.Drawing.Point(266, 78);
            this.txtLOADate.Name = "txtLOADate";
            this.txtLOADate.Size = new System.Drawing.Size(155, 26);
            this.txtLOADate.TabIndex = 0;
            // 
            // Startbutton
            // 
            this.Startbutton.BackColor = System.Drawing.Color.LightGray;
            this.Startbutton.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.Startbutton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.Startbutton.Location = new System.Drawing.Point(101, 150);
            this.Startbutton.Margin = new System.Windows.Forms.Padding(4);
            this.Startbutton.Name = "Startbutton";
            this.Startbutton.Size = new System.Drawing.Size(128, 38);
            this.Startbutton.TabIndex = 1;
            this.Startbutton.Text = "Go";
            this.toolTip1.SetToolTip(this.Startbutton, "Populates Excel File with Data From AS400");
            this.Startbutton.UseVisualStyleBackColor = false;
            this.Startbutton.Click += new System.EventHandler(this.Startbutton_Click);
            // 
            // CloseButton
            // 
            this.CloseButton.BackColor = System.Drawing.Color.LightGray;
            this.CloseButton.FlatStyle = System.Windows.Forms.FlatStyle.Popup;
            this.CloseButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.CloseButton.Location = new System.Drawing.Point(287, 150);
            this.CloseButton.Name = "CloseButton";
            this.CloseButton.Size = new System.Drawing.Size(134, 38);
            this.CloseButton.TabIndex = 2;
            this.CloseButton.Text = "Exit";
            this.toolTip1.SetToolTip(this.CloseButton, "Log-Off and End Program ");
            this.CloseButton.UseVisualStyleBackColor = false;
            this.CloseButton.Click += new System.EventHandler(this.CloseButton_Click);
            // 
            // ProgressBar1
            // 
            this.ProgressBar1.Location = new System.Drawing.Point(1, 285);
            this.ProgressBar1.Name = "ProgressBar1";
            this.ProgressBar1.Size = new System.Drawing.Size(490, 23);
            this.ProgressBar1.Step = 1;
            this.ProgressBar1.TabIndex = 1;
            this.ProgressBar1.Visible = false;
            // 
            // lblprogress1
            // 
            this.lblprogress1.AutoSize = true;
            this.lblprogress1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold);
            this.lblprogress1.ForeColor = System.Drawing.Color.MidnightBlue;
            this.lblprogress1.Location = new System.Drawing.Point(412, 262);
            this.lblprogress1.Name = "lblprogress1";
            this.lblprogress1.Size = new System.Drawing.Size(0, 20);
            this.lblprogress1.TabIndex = 36;
            this.lblprogress1.Visible = false;
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(494, 308);
            this.Controls.Add(this.lblprogress1);
            this.Controls.Add(this.ProgressBar1);
            this.Controls.Add(this.CloseButton);
            this.Controls.Add(this.Startbutton);
            this.Controls.Add(this.txtLOADate);
            this.Controls.Add(this.lblLOADate);
            this.Name = "frmMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "IP Renewal Utility";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        internal System.Windows.Forms.Label lblLOADate;
        internal System.Windows.Forms.DateTimePicker txtLOADate;
        internal System.Windows.Forms.Button Startbutton;
        internal System.Windows.Forms.Button CloseButton;
        internal System.Windows.Forms.ProgressBar ProgressBar1;
        internal System.Windows.Forms.ToolTip toolTip1;
        internal System.Windows.Forms.Label lblprogress1;
    }
}