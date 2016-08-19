namespace IP_NZ
{
    partial class frmService
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
            this.clearList = new System.Windows.Forms.Button();
            this.RenewalLabel = new System.Windows.Forms.Label();
            this.validateList = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // clearList
            // 
            this.clearList.Location = new System.Drawing.Point(34, 216);
            this.clearList.Name = "clearList";
            this.clearList.Size = new System.Drawing.Size(109, 56);
            this.clearList.TabIndex = 0;
            this.clearList.Text = "Clear List";
            this.clearList.UseVisualStyleBackColor = true;
            this.clearList.Click += new System.EventHandler(this.clearList_Click);
            // 
            // RenewalLabel
            // 
            this.RenewalLabel.Location = new System.Drawing.Point(31, 9);
            this.RenewalLabel.Name = "RenewalLabel";
            this.RenewalLabel.Size = new System.Drawing.Size(265, 98);
            this.RenewalLabel.TabIndex = 1;
            this.RenewalLabel.Text = "IPONZ Multiple IP Renewal Service";
            this.RenewalLabel.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.RenewalLabel.Click += new System.EventHandler(this.RenewalLabel_Click);
            // 
            // validateList
            // 
            this.validateList.Location = new System.Drawing.Point(171, 216);
            this.validateList.Name = "validateList";
            this.validateList.Size = new System.Drawing.Size(112, 56);
            this.validateList.TabIndex = 0;
            this.validateList.Text = "Validate List";
            this.validateList.UseVisualStyleBackColor = true;
            this.validateList.Click += new System.EventHandler(this.validateList_Click);
            // 
            // frmService
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(320, 300);
            this.Controls.Add(this.validateList);
            this.Controls.Add(this.RenewalLabel);
            this.Controls.Add(this.clearList);
            this.Name = "frmService";
            this.Text = "frmService";
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button clearList;
        private System.Windows.Forms.Label RenewalLabel;
        private System.Windows.Forms.Button validateList;
    }
}