using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using LinqToExcel;
using Microsoft.VisualBasic;
using PdfSearchHighlight.MyNSpace;
using PdfTextHighlighter.Code;


namespace PdfTextHighlighter
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void btnStart_Click(object sender, EventArgs e)
        {
            
            //if (string.IsNullOrEmpty(txtExcelFile.Text) || string.IsNullOrEmpty(txtFirstPDF.Text) ||
            //    string.IsNullOrEmpty(txtSecondPDF.Text))
            //{
            //    MessageBox.Show("One or more files not specified.");
            //    return;
            //}


            var pathToExcelFile = txtExcelFile.Text;

            const string sheetName = "Sheet1";

            var excelFile = new ExcelQueryFactory(pathToExcelFile);
            var artistAlbums = from a in excelFile.WorksheetNoHeader(sheetName) select a;

            StringBuilder sbArr = new StringBuilder();


            foreach (var a in artistAlbums)
            {
                sbArr.Append(a[1].Value.ToString());

            }


            PdfTextGetter(StringComparison.Ordinal, "b.pdf", "b_001.pdf", sbArr.ToString());
          
        }

        public void PdfTextGetter( StringComparison sc, string sourceFile, string destinationFile,string snewSearch)
        {

            string[] sArr = snewSearch.Split(',');

            this.Cursor = Cursors.WaitCursor;
            if (File.Exists(sourceFile))
            {
                PdfReader pReader = new PdfReader(sourceFile);


                var stamper = new iTextSharp.text.pdf.PdfStamper(pReader, new FileStream(destinationFile, FileMode.Append));
               
                
                PB.Value = 0;
                PB.Maximum = sArr.Length;

                foreach (var s in sArr)
                {
                    
                
                for (int page = 1; page <= pReader.NumberOfPages; page++)
                {

                    var t = new MyLocationTextExtractionStrategy(s,CompareOptions.Ordinal);


                    using (var r = new PdfReader(sourceFile))
                    {
                        var ex = PdfTextExtractor.GetTextFromPage(r, 1, t);
                    }

                    myLocationTextExtractionStrategy strategy = new myLocationTextExtractionStrategy();
                    var cb = stamper.GetUnderContent(page);

                    //Send some data contained in PdfContentByte, looks like the first is always cero for me and the second 100, but i'm not sure if this could change in some cases
                    strategy.UndercontentCharacterSpacing = (int) cb.CharacterSpacing;
                    strategy.UndercontentHorizontalScaling = (int) cb.HorizontalScaling;


                    //It's not really needed to get the text back, but we have to call this line ALWAYS, 
                    //because it triggers the process that will get all chunks from PDF into our strategy Object
                    string currentText = PdfTextExtractor.GetTextFromPage(pReader, page, strategy);

                    //The real getter process starts in the following line
                    List<RectAndText> matchesFound = t.myPoints;

                    //Set the fill color of the shapes, I don't use a border because it would make the rect bigger
                    //but maybe using a thin border could be a solution if you see the currect rect is not big enough to cover all the text it should cover
                    cb.SetColorFill(BaseColor.BLACK);
                    //.PINK)

                    //MatchesFound contains all text with locations, so do whatever you want with it, this highlights them using PINK color:

                    foreach (RectAndText rect in matchesFound)
                    {
                        if (rect.Text == s)
                        {
                            cb.Rectangle(rect.Rect.Left, rect.Rect.Bottom, rect.Rect.Width, rect.Rect.Height);
                            //Create our hightlight
                            float[] quad =
                            {
                                rect.Rect.Left,
                                rect.Rect.Bottom,
                                rect.Rect.Right,
                                rect.Rect.Bottom,
                                rect.Rect.Left,
                                rect.Rect.Top,
                                rect.Rect.Right,
                                rect.Rect.Top
                            };

                            PdfAnnotation highlight = PdfAnnotation.CreateMarkup(stamper.Writer, rect.Rect,
                                Constants.vbNull.ToString(), PdfAnnotation.MARKUP_HIGHLIGHT, quad);
                            //Set the color
                            highlight.Color = BaseColor.YELLOW;
                            //Add the annotation
                            stamper.AddAnnotation(highlight, page);
                        }
                    }
                    
                    PB.Value = PB.Value + 1;
                }
                }
                stamper.Close();
            }
            this.Cursor = Cursors.Default;

        }

        private void btnExcelFile_Click(object sender, EventArgs e)
        {
            var openFileDialogExcel = new OpenFileDialog
            {
                InitialDirectory = @"D:\Projects\Adrian P\test_files\",
                Filter = @"Excel Files|*.xls;*.xlsx",
                FilterIndex = 2,
                RestoreDirectory = true
            };

            if (openFileDialogExcel.ShowDialog() == DialogResult.OK)
            {
                txtExcelFile.Text = openFileDialogExcel.FileName;
            }
        }

        private void btnFirstPDF_Click(object sender, EventArgs e)
        {
            var openFileDialogPdf = new OpenFileDialog
            {
                InitialDirectory = @"D:\Projects\Adrian P\test_files\",
                Filter = @"Pdf Files|*.pdf",
                RestoreDirectory = true
            };

            if (openFileDialogPdf.ShowDialog() == DialogResult.OK)
            {
                txtFirstPDF.Text = openFileDialogPdf.FileName;
            }
        }

        private void btnSecondPDF_Click(object sender, EventArgs e)
        {
            var openFileDialogPdf = new OpenFileDialog
            {
                InitialDirectory = @"D:\Projects\Adrian P\test_files\",
                Filter = @"Pdf Files|*.pdf",
                RestoreDirectory = true
            };

            if (openFileDialogPdf.ShowDialog() == DialogResult.OK)
            {
                txtSecondPDF.Text = openFileDialogPdf.FileName;
            }
        }

        
    }
}
