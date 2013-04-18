using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using iTextSharp.text.pdf.parser;
using Application = System.Windows.Forms.Application;
using System.Runtime.InteropServices;
//using iTextSharp;
using iTextSharp.text.pdf;







namespace opening_word_document
{
    public partial class MainForm : Form
    {
        private OpenFileDialog ofd;
        private string pathToPdf;
      

        public MainForm()
        {
            InitializeComponent();
            ofd = new OpenFileDialog();
            
        }

        
        
        private void ConvertButton_Click(object sender, EventArgs e)
        {



            
            
            
            //POS Tagging code
            POSTagged post = new POSTagged();

            POSTagger.mModelPath = "Models\\";

            

            string content = DocText.Text;

            string[] tokenize = POSTagger.TokenizeSentence(content);
            string[] POS = POSTagger.PosTagTokens(tokenize);
            string POSTextbox = string.Empty;
            for (int i = 0; i < POS.Length; i++)
            {
                //NN Noun, singular or mass
                //NNS Noun, plural
                //NNP Proper noun, singular
                //NNPS Proper noun, plural
                if (POS[i] == "NN" || POS[i] == "NNS" || POS[i] == "NNP" || POS[i] == "NNPS")
                    POSTextbox = POSTextbox + (tokenize[i] + "/" + POS[i] + "  ");


            }
            post.PosTaggedText = POSTextbox;

            this.Hide();

            post.ShowDialog();

          
            this.Show();
            
        }

        private void BrowseBtn_Click(object sender, EventArgs e)
        {
            
            ofd.Filter = "Word Document(*.doc) | *.doc|PDF(*.pdf)|*.pdf|Open Doc Text(*.odt)|*.odt|Microsoft XPS(*.xps)|*.xps";
            
            
            if (ofd.ShowDialog() == DialogResult.OK)
            textPathName.Text = ofd.FileName;
            
            
        }

        private void ReadButton_Click(object sender, EventArgs e)
        {
            if (textPathName.Text.Length > 0)
            {
                ReadFileContent(textPathName.Text);
            }
            else
            {
                MessageBox.Show("Enter a valid file path");
            }
        }

        public void ReadFileContent(string path)
        {
            string ext = Path.GetExtension(path);
            if (ext == ".doc")
            {

                try
                {
                    //var wordApp = new ApplicationClass();
                    //object file = path;
                    //object nullobj = System.Reflection.Missing.Value;

                    //Document doc = wordApp.Documents.Open(
                    //    ref file, ref nullobj, ref nullobj,
                    //    ref nullobj, ref nullobj, ref nullobj,
                    //    ref nullobj, ref nullobj, ref nullobj,
                    //    ref nullobj, ref nullobj, ref nullobj);
                    //doc.ActiveWindow.Selection.WholeStory();
                    //doc.ActiveWindow.Selection.Copy();
                    //IDataObject data = Clipboard.GetDataObject();
                    //DocText.Text = data.GetData(DataFormats.Text).ToString();
                    //doc.Close(ref nullobj, ref nullobj, ref nullobj);
                    //wordApp.Quit(ref nullobj, ref nullobj, ref nullobj);

                    Word2pdf w2p=new Word2pdf();
                    pathToPdf= w2p.ConvertToPdf(path);
                    ReadPdf(pathToPdf);

                }

                catch (COMException)
                {
                    MessageBox.Show("Unable to read this document.  It may be corrupt.");

                }
            }

            else
            {
                ReadPdf(path);
            }

        }

        public void ReadPdf(string path)
        {
            try
            {
                MessageBox.Show("starting to read pdf");
                PdfReader pdfr = new PdfReader(path);
                StringBuilder pdfText = new StringBuilder();

                for (int page = 1; page <= pdfr.NumberOfPages; page++)
                {
                    ITextExtractionStrategy strategy = new SimpleTextExtractionStrategy();
                    string currentText = PdfTextExtractor.GetTextFromPage(pdfr, page, strategy);

                    currentText = Encoding.UTF8.GetString(ASCIIEncoding.Convert(Encoding.Default, Encoding.UTF8, Encoding.Default.GetBytes(currentText)));
                    pdfText.Append(currentText);
                }

                pdfr.Close();
                DocText.Text = pdfText.ToString();
            }
            catch (Exception)
            {

                MessageBox.Show("problem with pdf file");
            }
        }

       
        
    }
}
