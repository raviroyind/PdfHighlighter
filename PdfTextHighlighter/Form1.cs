using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using LinqToExcel;
using Microsoft.VisualBasic;
using PdfTextHighlighter.Code;


namespace PdfTextHighlighter
{
    public partial class Form1 : Form
    {
        List<string> _fileList = new List<string>();
        private string _currentFile = string.Empty;
        private string _lastBrowseLocation = string.Empty;
        public Form1()
        {
            InitializeComponent();
        }

        #region Button Events...

        private void btnStart_Click(object sender, EventArgs e)
        {
            _fileList=new List<string>();

            if (string.IsNullOrEmpty(txtExcelFile.Text) || string.IsNullOrEmpty(txtFirstPDF.Text) ||
                string.IsNullOrEmpty(txtSecondPDF.Text) || string.IsNullOrEmpty(txtDestinationFolder.Text))
            {
                MessageBox.Show("One or more files not specified.");
                return;
            }


            var pathToExcelFile = txtExcelFile.Text;

            const string sheetName = "Sheet1";

            var excelFile = new ExcelQueryFactory(pathToExcelFile);
            var columnValues = from a in excelFile.WorksheetNoHeader(sheetName) select a;


            List<ExcelRow> listRows = new List<ExcelRow>();


            var setNumber = 1;
            var increment = 0;
            foreach (var a in columnValues)
            {
                listRows.Add(
                    new ExcelRow
                    {
                        Index = increment,
                        SetNumber = setNumber,
                        ColumnValue = a[1].Value.ToString(),
                        FileName = string.IsNullOrWhiteSpace(a[0].Value.ToString())?"":a[0].Value.ToString()
                    }
                );


                if (!a[1].Value.ToString().EndsWith(","))
                    setNumber++;

                increment++;
            }

            for (var i = 1; i < increment - 1; i++)
            {
                var sbSearch = new StringBuilder();

                foreach (var item in listRows.Where(item => item.SetNumber == i))
                {
                    sbSearch.Append(item.ColumnValue);
                    if (!string.IsNullOrEmpty(item.FileName))
                        _currentFile = item.FileName;
                }

               

               if(!string.IsNullOrEmpty(sbSearch.ToString()))
                   ProcessPdf(StringComparison.Ordinal, txtFirstPDF.Text, txtDestinationFolder.Text + "\\" + GetFileName(_currentFile, txtFirstPDF.Text) + ".pdf", sbSearch.ToString());


               if (!string.IsNullOrEmpty(txtSecondPDF.Text) && !string.IsNullOrEmpty(sbSearch.ToString()))
                   ProcessPdf(StringComparison.Ordinal, txtSecondPDF.Text, txtDestinationFolder.Text + "\\" + GetFileName(_currentFile, txtSecondPDF.Text) + ".pdf", sbSearch.ToString());
                    
            }

            

            MessageBox.Show("Pdf highlighted successfully!");

            if (chkOpenPdfs.Checked)
            {

                if (_fileList.Count != 0)
                {
                    foreach (var file in _fileList)
                    {
                        System.Diagnostics.Process.Start(file);
                    }
                }
            }

        }

        private static string GetFileName(string currentFile,string sourceFile)
        {
            var paddingLeft = string.Empty;

            if (currentFile.Length==1)
                paddingLeft =  "_00";
            else if (currentFile.Length == 2)
                paddingLeft = "_0";
            else if (currentFile.Length == 3)
                paddingLeft = "_";

            return System.IO.Path.GetFileNameWithoutExtension(sourceFile) + paddingLeft + currentFile;
        }

        private void btnExcelFile_Click(object sender, EventArgs e)
        {
            var openFileDialogExcel = new OpenFileDialog
            {
                InitialDirectory = string.IsNullOrEmpty(_lastBrowseLocation) ? @"C:\" : _lastBrowseLocation,
                Filter = @"Excel Files|*.xls;*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = false
                
            };

            if (openFileDialogExcel.ShowDialog() == DialogResult.OK)
            {
                _lastBrowseLocation = openFileDialogExcel.FileName.Substring(0,
                    openFileDialogExcel.FileName.LastIndexOf("\\", StringComparison.Ordinal));

                txtExcelFile.Text = openFileDialogExcel.FileName;
            }
        }

        private void btnFirstPDF_Click(object sender, EventArgs e)
        {
            var openFileDialogPdf = new OpenFileDialog
            {
                InitialDirectory = string.IsNullOrEmpty(_lastBrowseLocation) ? @"C:\" : _lastBrowseLocation,
                Filter = @"Pdf Files|*.pdf",
                RestoreDirectory = false
            };

            if (openFileDialogPdf.ShowDialog() == DialogResult.OK)
            {
                _lastBrowseLocation = openFileDialogPdf.FileName.Substring(0,
                    openFileDialogPdf.FileName.LastIndexOf("\\", StringComparison.Ordinal));

                txtFirstPDF.Text = openFileDialogPdf.FileName;
            }
        }

        private void btnSecondPDF_Click(object sender, EventArgs e)
        {
            var openFileDialogPdf = new OpenFileDialog
            {
                InitialDirectory = string.IsNullOrEmpty(_lastBrowseLocation) ? @"C:\" : _lastBrowseLocation,
                Filter = @"Pdf Files|*.pdf",
                RestoreDirectory = false
            };

            if (openFileDialogPdf.ShowDialog() == DialogResult.OK)
            {
                _lastBrowseLocation = openFileDialogPdf.FileName.Substring(0,
                   openFileDialogPdf.FileName.LastIndexOf("\\", StringComparison.Ordinal));

                txtSecondPDF.Text = openFileDialogPdf.FileName;
            }
        }

        private void chkOpenPdfs_CheckedChanged(object sender, EventArgs e)
        {
            if(chkOpenPdfs.Checked)
                MessageBox.Show(this,@"Checking this will open all the resulting Pdfs!", @"Warming",MessageBoxButtons.OK);
        }

        private void btnDestinationFolder_Click(object sender, EventArgs e)
        {
            var openFolderDialog = new FolderBrowserDialog
            {
                RootFolder = Environment.SpecialFolder.Desktop,
                ShowNewFolderButton = true,

            };

            if (openFolderDialog.ShowDialog() == DialogResult.OK)
            {
                txtDestinationFolder.Text = openFolderDialog.SelectedPath;
            }
        }

#endregion Button Events...

#region Highlight...

        public void ProcessPdf(StringComparison sc, string sourceFile, string destinationFile, string searchTerm)
        {

            var sArr = searchTerm.Split(',');
            var exists = false;
         
            foreach (var occurence in sArr.Select(item => ReadPdfFile(sourceFile, item)).Where(occurence => occurence.Count > 0))
            {
                exists = true;
            }

            if (!exists)
                return;

            Cursor = Cursors.WaitCursor;
            if (File.Exists(sourceFile))
            {
                var pReader = new PdfReader(sourceFile);

                var stamper = new PdfStamper(pReader, new FileStream(destinationFile, FileMode.Append));

                _fileList.Add(destinationFile);

                progressBar.Value = 0;
                progressBar.Maximum = sArr.Length;
                
                foreach (var s in sArr)
                {


                    for (var page = 1; page <= pReader.NumberOfPages; page++)
                    {

                        var t = new MyLocationTextExtractionStrategy(s, CompareOptions.Ordinal);


                        using (var r = new PdfReader(sourceFile))
                        {
                            var ex = PdfTextExtractor.GetTextFromPage(r, 1, t);
                        }


                        var cb = stamper.GetUnderContent(page);

                        var matchesFound = t.MyPoints;

                        
                        cb.SetColorFill(BaseColor.BLACK);
                        

                        foreach (var rect in matchesFound)
                        {
                            if (rect.Text == s)
                            {
                                cb.Rectangle(rect.Rect.Left, rect.Rect.Bottom, rect.Rect.Width, rect.Rect.Height);

                                float[] quad = {
                                    rect.Rect.Left,
                                    rect.Rect.Bottom,
                                    rect.Rect.Right,
                                    rect.Rect.Bottom,
                                    rect.Rect.Left,
                                    rect.Rect.Top,
                                    rect.Rect.Right,
                                    rect.Rect.Top
                                };


                                var highlight = PdfAnnotation.CreateMarkup(stamper.Writer, rect.Rect,
                                    Constants.vbNull.ToString(), PdfAnnotation.MARKUP_HIGHLIGHT, quad);
                                
                                highlight.Color = BaseColor.YELLOW;

                                stamper.AddAnnotation(highlight, page);
                            }
                        }

                        progressBar.Value = progressBar.Value + 1;
                    }
                }
                stamper.Close();
            }
            this.Cursor = Cursors.Default;

        }


        public List<int> ReadPdfFile(string fileName, string searthText)
        {
            List<int> pages = new List<int>();
            if (File.Exists(fileName))
            {
                PdfReader pdfReader = new PdfReader(fileName);
                for (int page = 1; page <= pdfReader.NumberOfPages; page++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();

                    string currentPageText = PdfTextExtractor.GetTextFromPage(pdfReader, page, strategy);
                    if (currentPageText.Contains(searthText))
                    {
                        pages.Add(page);
                    }
                }
                pdfReader.Close();
            }
            return pages;
        }
#endregion Highlight...

    
    }
}
