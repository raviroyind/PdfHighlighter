namespace PdfTextHighlighter
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
            this.btnStart = new System.Windows.Forms.Button();
            this.PB = new System.Windows.Forms.ProgressBar();
            this.btnExcelFile = new System.Windows.Forms.Button();
            this.txtExcelFile = new System.Windows.Forms.TextBox();
            this.txtFirstPDF = new System.Windows.Forms.TextBox();
            this.btnFirstPDF = new System.Windows.Forms.Button();
            this.txtSecondPDF = new System.Windows.Forms.TextBox();
            this.btnSecondPDF = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(340, 299);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(233, 105);
            this.btnStart.TabIndex = 0;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // PB
            // 
            this.PB.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.PB.Location = new System.Drawing.Point(26, 196);
            this.PB.Margin = new System.Windows.Forms.Padding(4);
            this.PB.Name = "PB";
            this.PB.Size = new System.Drawing.Size(547, 73);
            this.PB.TabIndex = 2;
            // 
            // btnExcelFile
            // 
            this.btnExcelFile.Location = new System.Drawing.Point(26, 24);
            this.btnExcelFile.Name = "btnExcelFile";
            this.btnExcelFile.Size = new System.Drawing.Size(145, 40);
            this.btnExcelFile.TabIndex = 4;
            this.btnExcelFile.Text = "Select Excel File";
            this.btnExcelFile.UseVisualStyleBackColor = true;
            this.btnExcelFile.Click += new System.EventHandler(this.btnExcelFile_Click);
            // 
            // txtExcelFile
            // 
            this.txtExcelFile.Location = new System.Drawing.Point(177, 33);
            this.txtExcelFile.Name = "txtExcelFile";
            this.txtExcelFile.Size = new System.Drawing.Size(396, 22);
            this.txtExcelFile.TabIndex = 5;
            // 
            // txtFirstPDF
            // 
            this.txtFirstPDF.Location = new System.Drawing.Point(177, 88);
            this.txtFirstPDF.Name = "txtFirstPDF";
            this.txtFirstPDF.Size = new System.Drawing.Size(396, 22);
            this.txtFirstPDF.TabIndex = 7;
            // 
            // btnFirstPDF
            // 
            this.btnFirstPDF.Location = new System.Drawing.Point(26, 79);
            this.btnFirstPDF.Name = "btnFirstPDF";
            this.btnFirstPDF.Size = new System.Drawing.Size(145, 40);
            this.btnFirstPDF.TabIndex = 6;
            this.btnFirstPDF.Text = "Select Pdf 1";
            this.btnFirstPDF.UseVisualStyleBackColor = true;
            this.btnFirstPDF.Click += new System.EventHandler(this.btnFirstPDF_Click);
            // 
            // txtSecondPDF
            // 
            this.txtSecondPDF.Location = new System.Drawing.Point(177, 148);
            this.txtSecondPDF.Name = "txtSecondPDF";
            this.txtSecondPDF.Size = new System.Drawing.Size(396, 22);
            this.txtSecondPDF.TabIndex = 9;
            // 
            // btnSecondPDF
            // 
            this.btnSecondPDF.Location = new System.Drawing.Point(26, 139);
            this.btnSecondPDF.Name = "btnSecondPDF";
            this.btnSecondPDF.Size = new System.Drawing.Size(145, 40);
            this.btnSecondPDF.TabIndex = 8;
            this.btnSecondPDF.Text = "Select Pdf 2";
            this.btnSecondPDF.UseVisualStyleBackColor = true;
            this.btnSecondPDF.Click += new System.EventHandler(this.btnSecondPDF_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1219, 735);
            this.Controls.Add(this.txtSecondPDF);
            this.Controls.Add(this.btnSecondPDF);
            this.Controls.Add(this.txtFirstPDF);
            this.Controls.Add(this.btnFirstPDF);
            this.Controls.Add(this.txtExcelFile);
            this.Controls.Add(this.btnExcelFile);
            this.Controls.Add(this.PB);
            this.Controls.Add(this.btnStart);
            this.Name = "Form1";
            this.Text = "Pdf Highlighter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnStart;
        internal System.Windows.Forms.ProgressBar PB;
        private System.Windows.Forms.Button btnExcelFile;
        private System.Windows.Forms.TextBox txtExcelFile;
        private System.Windows.Forms.TextBox txtFirstPDF;
        private System.Windows.Forms.Button btnFirstPDF;
        private System.Windows.Forms.TextBox txtSecondPDF;
        private System.Windows.Forms.Button btnSecondPDF;
    }
}

