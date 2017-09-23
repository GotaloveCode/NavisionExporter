namespace ayoti
{
    partial class Form1
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
            this.button1 = new System.Windows.Forms.Button();
            this.lblstatus = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtworksheet = new System.Windows.Forms.TextBox();
            this.txtrows = new System.Windows.Forms.TextBox();
            this.btnview = new System.Windows.Forms.Button();
            this.txtstatus = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.txtworksheetcount = new System.Windows.Forms.TextBox();
            this.btnclear = new System.Windows.Forms.Button();
            this.btnGetExcel = new System.Windows.Forms.Button();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.chkoverride = new System.Windows.Forms.CheckBox();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.lblfile = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Enabled = false;
            this.button1.Location = new System.Drawing.Point(42, 55);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(122, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Import";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // lblstatus
            // 
            this.lblstatus.AutoSize = true;
            this.lblstatus.Location = new System.Drawing.Point(39, 176);
            this.lblstatus.Name = "lblstatus";
            this.lblstatus.Size = new System.Drawing.Size(37, 13);
            this.lblstatus.TabIndex = 1;
            this.lblstatus.Text = "Status";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(39, 206);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Rows Affected:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(194, 89);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(92, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Select Worksheet";
            // 
            // txtworksheet
            // 
            this.txtworksheet.Location = new System.Drawing.Point(197, 105);
            this.txtworksheet.Name = "txtworksheet";
            this.txtworksheet.ReadOnly = true;
            this.txtworksheet.Size = new System.Drawing.Size(57, 20);
            this.txtworksheet.TabIndex = 4;
            // 
            // txtrows
            // 
            this.txtrows.Location = new System.Drawing.Point(197, 206);
            this.txtrows.Name = "txtrows";
            this.txtrows.ReadOnly = true;
            this.txtrows.Size = new System.Drawing.Size(100, 20);
            this.txtrows.TabIndex = 5;
            // 
            // btnview
            // 
            this.btnview.Location = new System.Drawing.Point(44, 100);
            this.btnview.Name = "btnview";
            this.btnview.Size = new System.Drawing.Size(122, 23);
            this.btnview.TabIndex = 6;
            this.btnview.Text = "View Imports";
            this.btnview.UseVisualStyleBackColor = true;
            this.btnview.Click += new System.EventHandler(this.btnview_Click);
            // 
            // txtstatus
            // 
            this.txtstatus.BackColor = System.Drawing.Color.WhiteSmoke;
            this.txtstatus.ForeColor = System.Drawing.Color.Red;
            this.txtstatus.Location = new System.Drawing.Point(197, 176);
            this.txtstatus.Name = "txtstatus";
            this.txtstatus.ReadOnly = true;
            this.txtstatus.Size = new System.Drawing.Size(100, 20);
            this.txtstatus.TabIndex = 7;
            this.txtstatus.Text = "Not started";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(267, 105);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(19, 13);
            this.label3.TabIndex = 8;
            this.label3.Text = "of ";
            // 
            // txtworksheetcount
            // 
            this.txtworksheetcount.Location = new System.Drawing.Point(292, 105);
            this.txtworksheetcount.Name = "txtworksheetcount";
            this.txtworksheetcount.ReadOnly = true;
            this.txtworksheetcount.Size = new System.Drawing.Size(72, 20);
            this.txtworksheetcount.TabIndex = 9;
            // 
            // btnclear
            // 
            this.btnclear.Location = new System.Drawing.Point(44, 139);
            this.btnclear.Name = "btnclear";
            this.btnclear.Size = new System.Drawing.Size(122, 23);
            this.btnclear.TabIndex = 10;
            this.btnclear.Text = "Clear Table";
            this.btnclear.UseVisualStyleBackColor = true;
            this.btnclear.Click += new System.EventHandler(this.btnclear_Click);
            // 
            // btnGetExcel
            // 
            this.btnGetExcel.Location = new System.Drawing.Point(44, 12);
            this.btnGetExcel.Name = "btnGetExcel";
            this.btnGetExcel.Size = new System.Drawing.Size(120, 23);
            this.btnGetExcel.TabIndex = 11;
            this.btnGetExcel.Text = "Get Excel";
            this.btnGetExcel.UseVisualStyleBackColor = true;
            this.btnGetExcel.Click += new System.EventHandler(this.btnGetExcel_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // chkoverride
            // 
            this.chkoverride.AutoSize = true;
            this.chkoverride.Location = new System.Drawing.Point(197, 55);
            this.chkoverride.Name = "chkoverride";
            this.chkoverride.Size = new System.Drawing.Size(116, 17);
            this.chkoverride.TabIndex = 12;
            this.chkoverride.Text = "override worksheet";
            this.chkoverride.UseVisualStyleBackColor = true;
            this.chkoverride.CheckedChanged += new System.EventHandler(this.chkoverride_CheckedChanged);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(12, 280);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(537, 96);
            this.richTextBox1.TabIndex = 14;
            this.richTextBox1.Text = "";
            // 
            // lblfile
            // 
            this.lblfile.AutoSize = true;
            this.lblfile.Location = new System.Drawing.Point(196, 17);
            this.lblfile.Name = "lblfile";
            this.lblfile.Size = new System.Drawing.Size(72, 13);
            this.lblfile.TabIndex = 15;
            this.lblfile.Text = "excel file path";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(561, 388);
            this.Controls.Add(this.lblfile);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.chkoverride);
            this.Controls.Add(this.btnGetExcel);
            this.Controls.Add(this.btnclear);
            this.Controls.Add(this.txtworksheetcount);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtstatus);
            this.Controls.Add(this.btnview);
            this.Controls.Add(this.txtrows);
            this.Controls.Add(this.txtworksheet);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.lblstatus);
            this.Controls.Add(this.button1);
            this.Name = "Form1";
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label lblstatus;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtrows;
        private System.Windows.Forms.Button btnview;
        private System.Windows.Forms.TextBox txtstatus;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox txtworksheetcount;
        private System.Windows.Forms.Button btnclear;
        private System.Windows.Forms.Button btnGetExcel;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.TextBox txtworksheet;
        private System.Windows.Forms.CheckBox chkoverride;
        private System.Windows.Forms.RichTextBox richTextBox1;
        private System.Windows.Forms.Label lblfile;
    }
}

