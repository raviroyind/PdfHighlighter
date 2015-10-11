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
            this.progressBar = new System.Windows.Forms.ProgressBar();
            this.btnExcelFile = new System.Windows.Forms.Button();
            this.txtExcelFile = new System.Windows.Forms.TextBox();
            this.txtFirstPDF = new System.Windows.Forms.TextBox();
            this.btnFirstPDF = new System.Windows.Forms.Button();
            this.txtSecondPDF = new System.Windows.Forms.TextBox();
            this.btnSecondPDF = new System.Windows.Forms.Button();
            this.txtDestinationFolder = new System.Windows.Forms.TextBox();
            this.btnDestinationFolder = new System.Windows.Forms.Button();
            this.chkOpenPdfs = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // btnStart
            // 
            this.btnStart.Location = new System.Drawing.Point(340, 436);
            this.btnStart.Name = "btnStart";
            this.btnStart.Size = new System.Drawing.Size(233, 55);
            this.btnStart.TabIndex = 0;
            this.btnStart.Text = "Start";
            this.btnStart.UseVisualStyleBackColor = true;
            this.btnStart.Click += new System.EventHandler(this.btnStart_Click);
            // 
            // progressBar
            // 
            this.progressBar.BackColor = System.Drawing.Color.FromArgb(((int)(((byte)(255)))), ((int)(((byte)(255)))), ((int)(((byte)(192)))));
            this.progressBar.Location = new System.Drawing.Point(26, 203);
            this.progressBar.Margin = new System.Windows.Forms.Padding(4);
            this.progressBar.Name = "progressBar";
            this.progressBar.Size = new System.Drawing.Size(547, 73);
            this.progressBar.TabIndex = 2;
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
            // txtDestinationFolder
            // 
            this.txtDestinationFolder.Location = new System.Drawing.Point(177, 347);
            this.txtDestinationFolder.Name = "txtDestinationFolder";
            this.txtDestinationFolder.Size = new System.Drawing.Size(396, 22);
            this.txtDestinationFolder.TabIndex = 11;
            // 
            // btnDestinationFolder
            // 
            this.btnDestinationFolder.Location = new System.Drawing.Point(26, 338);
            this.btnDestinationFolder.Name = "btnDestinationFolder";
            this.btnDestinationFolder.Size = new System.Drawing.Size(145, 40);
            this.btnDestinationFolder.TabIndex = 10;
            this.btnDestinationFolder.Text = "Destination Folder";
            this.btnDestinationFolder.UseVisualStyleBackColor = true;
            this.btnDestinationFolder.Click += new System.EventHandler(this.btnDestinationFolder_Click);
            // 
            // chkOpenPdfs
            // 
            this.chkOpenPdfs.AutoSize = true;
            this.chkOpenPdfs.Location = new System.Drawing.Point(381, 393);
            this.chkOpenPdfs.Name = "chkOpenPdfs";
            this.chkOpenPdfs.Size = new System.Drawing.Size(192, 21);
            this.chkOpenPdfs.TabIndex = 12;
            this.chkOpenPdfs.Text = " Open Pdf(s) when done?";
            this.chkOpenPdfs.UseVisualStyleBackColor = true;
            this.chkOpenPdfs.CheckedChanged += new System.EventHandler(this.chkOpenPdfs_CheckedChanged);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.Control;
            this.ClientSize = new System.Drawing.Size(602, 547);
            this.Controls.Add(this.chkOpenPdfs);
            this.Controls.Add(this.txtDestinationFolder);
            this.Controls.Add(this.btnDestinationFolder);
            this.Controls.Add(this.txtSecondPDF);
            this.Controls.Add(this.btnSecondPDF);
            this.Controls.Add(this.txtFirstPDF);
            this.Controls.Add(this.btnFirstPDF);
            this.Controls.Add(this.txtExcelFile);
            this.Controls.Add(this.btnExcelFile);
            this.Controls.Add(this.progressBar);
            this.Controls.Add(this.btnStart);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Pdf Highlighter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnStart;
        internal System.Windows.Forms.ProgressBar progressBar;
        private System.Windows.Forms.Button btnExcelFile;
        private System.Windows.Forms.TextBox txtExcelFile;
        private System.Windows.Forms.TextBox txtFirstPDF;
        private System.Windows.Forms.Button btnFirstPDF;
        private System.Windows.Forms.TextBox txtSecondPDF;
        private System.Windows.Forms.Button btnSecondPDF;
        private System.Windows.Forms.TextBox txtDestinationFolder;
        private System.Windows.Forms.Button btnDestinationFolder;
        private System.Windows.Forms.CheckBox chkOpenPdfs;
    }
}

