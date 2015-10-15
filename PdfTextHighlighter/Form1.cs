using System;
using System.Collections.Generic;
using System.Data.OleDb;
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
using ExcelRow = PdfTextHighlighter.Code.ExcelRow;


namespace PdfTextHighlighter
{
    public partial class Form1 : Form
    {
        List<string> _fileList = new List<string>();
        private string _currentFile = string.Empty;
        private int _actualRow = 0;
        private string _lastBrowseLocation = string.Empty;
        //const string SheetName = "Sheet1";


        List<KeyValuePair<int, string>> _searchValues = new List<KeyValuePair<int, string>>();
        List<KeyValuePair<int, string>> _foundValuesInFirstPdf = new List<KeyValuePair<int, string>>();
        List<KeyValuePair<int, string>> _foundValuesInSecindPdf = new List<KeyValuePair<int, string>>();
        public Form1()
        {
            InitializeComponent();
            Application.EnableVisualStyles();
        }

        #region Button Events...

        private void btnStart_Click(object sender, EventArgs e)
        {
            CleanDestination();

            _fileList = new List<string>();
            btnStart.Enabled = false;

            if (string.IsNullOrEmpty(txtExcelFile.Text) || string.IsNullOrEmpty(txtFirstPDF.Text) ||
                string.IsNullOrEmpty(txtSecondPDF.Text) || string.IsNullOrEmpty(txtDestinationFolder.Text))
            {
                MessageBox.Show("One or more files not specified.");
                return;
            }


            var pathToExcelFile = txtExcelFile.Text;
            var excelFile = new ExcelQueryFactory(pathToExcelFile);
            var columnValues = from a in excelFile.WorksheetNoHeader(0) select a;


            var listRows = new List<ExcelRow>();


            var setNumber = 1;
            var increment = 1;
            foreach (var a in columnValues)
            {
                listRows.Add(
                    new ExcelRow
                    {
                        Index = increment,
                        SetNumber = setNumber,
                        ColumnValue = a[1].Value.ToString(),
                        FileName = string.IsNullOrWhiteSpace(a[0].Value.ToString()) ? "" : a[0].Value.ToString()
                    }
                );


                if (!a[1].Value.ToString().EndsWith(","))
                    setNumber++;

                increment++;
            }

            for (var i = 1; i < increment - 1; i++)
            {
                var sbSearch = new StringBuilder();
                _searchValues.Clear();

                foreach (var item in from item in listRows.Where(item => item.SetNumber == i) let index = item.Index let text = item.ColumnValue select item)
                {
                    sbSearch.Append(item.ColumnValue);

                    _searchValues.Add(new KeyValuePair<int, string>(item.Index, item.ColumnValue));

                    if (!string.IsNullOrEmpty(item.FileName))
                    {
                        _currentFile = item.FileName;
                        _actualRow = item.SetNumber;
                    }
                }

                if (!string.IsNullOrEmpty(sbSearch.ToString()))
                    ProcessPdf(StringComparison.Ordinal, txtFirstPDF.Text, txtDestinationFolder.Text + "\\" + GetFileName(_currentFile, txtFirstPDF.Text) + ".pdf", sbSearch.ToString(), i, _searchValues);


                if (!string.IsNullOrEmpty(txtSecondPDF.Text) && !string.IsNullOrEmpty(sbSearch.ToString()))
                    ProcessPdf(StringComparison.Ordinal, txtSecondPDF.Text, txtDestinationFolder.Text + "\\" + GetFileName(_currentFile, txtSecondPDF.Text) + ".pdf", sbSearch.ToString(), i, _searchValues);

            }

            #region Update Excel...

            lblMsg.Visible = true;
            foreach (var item in _foundValuesInFirstPdf)
            {
                var myCommand = new OleDbCommand();
                string sqlUpdate = null;
                var count = item.Value.Split(',').Count();

                var cnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + txtExcelFile.Text + ";Extended Properties='Excel 8.0;HDR=NO;'";
                var myConnection = new OleDbConnection(cnn);
                myConnection.Open();
                myCommand.Connection = myConnection;
                sqlUpdate = "UPDATE [Sheet1$H" + item.Key + ":H" + item.Key + "] SET F1='" + item.Value + "'";
                myCommand.CommandText = sqlUpdate;
                myCommand.ExecuteNonQuery();


                sqlUpdate = "UPDATE [Sheet1$J" + item.Key + ":J" + item.Key + "] SET F1='" + count + "'";
                myCommand.CommandText = sqlUpdate;
                myCommand.ExecuteNonQuery();

                myConnection.Close();

            }

            foreach (var item in _foundValuesInSecindPdf)
            {
                var myCommand = new OleDbCommand();
                string sqlUpdate = null;
                var count = item.Value.Split(',').Count();

                var cnn = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + txtExcelFile.Text + ";Extended Properties='Excel 8.0;HDR=NO;'";
                var myConnection = new OleDbConnection(cnn);
                myConnection.Open();
                myCommand.Connection = myConnection;
                sqlUpdate = "UPDATE [Sheet1$I" + item.Key + ":I" + item.Key + "] SET F1='" + item.Value + "'";
                myCommand.CommandText = sqlUpdate;
                myCommand.ExecuteNonQuery();

                sqlUpdate = "UPDATE [Sheet1$K" + item.Key + ":K" + item.Key + "] SET F1='" + count + "'";
                myCommand.CommandText = sqlUpdate;
                myCommand.ExecuteNonQuery();

                myConnection.Close();

            }

            #endregion Update Excel...

            MessageBox.Show("Pdf highlighted successfully!");
            lblMsg.Visible = false;
            btnStart.Enabled = true;

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

        private void CleanDestination()
        {
            var downloadedMessageInfo = new DirectoryInfo(txtDestinationFolder.Text);

            foreach (var file in downloadedMessageInfo.GetFiles())
            {
                try
                {
                    file.Delete();
                }
                catch (Exception)
                {
                    throw;
                }
                
            }
        }

        private static string GetFileName(string currentFile, string sourceFile)
        {

            var paddingLeft = string.Empty;

            if (currentFile.Length == 1)
                paddingLeft = "_00";
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
            if (chkOpenPdfs.Checked)
                MessageBox.Show(this, @"Checking this will open all the resulting Pdfs!", @"Warming", MessageBoxButtons.OK);
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

        public void ProcessPdf(StringComparison sc, string sourceFile, string destinationFile, string searchTerm, int excelRowNumber, List<KeyValuePair<int, string>> searchValues)
        {
            var sArr = searchTerm.Split(',');
            myProgressBar.Maximum = searchValues.Count;
            bool found = false;
            Cursor = Cursors.WaitCursor;
            if (File.Exists(sourceFile))
            {

                var pReader = new PdfReader(sourceFile);
                myProgressBar.Value = 0;
                PdfStamper stamper = null;

                foreach (var item in searchValues)
                {

                    var newStrings = item.Value.Split(',');
                    var foundText = string.Empty;
                    foreach (var search in newStrings)
                    {

                        for (var page = 1; page <= pReader.NumberOfPages; page++)
                        {

                            var t = new MyLocationTextExtractionStrategy(search, CompareOptions.Ordinal);

                            using (var r = new PdfReader(sourceFile))
                            {
                                var ex = PdfTextExtractor.GetTextFromPage(r, 1, t);
                            }

                            var matchesFound = t.MyPoints;

                            if (t.MyPoints.Count > 0)
                            {
                                found = true;
                                if (!string.IsNullOrEmpty(search))
                                {
                                    foundText += "," + search;
                                }


                                if (!File.Exists(destinationFile))
                                    stamper = new PdfStamper(pReader, new FileStream(destinationFile, FileMode.Create));

                                if (!_fileList.Contains(destinationFile))
                                    _fileList.Add(destinationFile);

                                var cb = stamper.GetUnderContent(page);
                                cb.SetColorFill(BaseColor.BLACK);

                                foreach (var rect in matchesFound)
                                {
                                    if (rect.Text == search)
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
                            }
                        }

                    }

                    if (found && !string.IsNullOrEmpty(foundText))
                    {

                        if (sourceFile == txtFirstPDF.Text)
                        {
                            _foundValuesInFirstPdf.Add(new KeyValuePair<int, string>(item.Key, foundText.Substring(1)));
                        }
                        else
                        {
                            _foundValuesInSecindPdf.Add(new KeyValuePair<int, string>(item.Key, foundText.Substring(1)));
                        }
                    }

                    myProgressBar.Value = myProgressBar.Value + 1;
                }

                if (stamper != null)
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
