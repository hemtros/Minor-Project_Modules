using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Office.Interop.Word;
using Application = System.Windows.Forms.Application;
using System.Runtime.InteropServices;






namespace opening_word_document
{
    public partial class MainForm : Form
    {
        
        public MainForm()
        {
            InitializeComponent();
            
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
                POSTextbox = POSTextbox + (tokenize[i] + "/" + POS[i] + "  ");


            }
            post.PosTaggedText = POSTextbox;

            this.Hide();

            post.ShowDialog();

          
            this.Show();
            
        }

        private void BrowseBtn_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd =new OpenFileDialog();
            ofd.Filter = "Word Document(*.doc) | *.doc";
            ofd.ShowDialog();
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
            try
            {
                var wordApp = new ApplicationClass();
                object file = path;
                object nullobj = System.Reflection.Missing.Value;

                Document doc = wordApp.Documents.Open(
                    ref file, ref nullobj, ref nullobj,
                    ref nullobj, ref nullobj, ref nullobj,
                    ref nullobj, ref nullobj, ref nullobj,
                    ref nullobj, ref nullobj, ref nullobj);
                doc.ActiveWindow.Selection.WholeStory();
                doc.ActiveWindow.Selection.Copy();
                IDataObject data = Clipboard.GetDataObject();
                DocText.Text = data.GetData(DataFormats.Text).ToString();
                doc.Close(ref nullobj, ref nullobj, ref nullobj);
                wordApp.Quit(ref nullobj, ref nullobj, ref nullobj);
            }

            catch(COMException)
            {
                MessageBox.Show("Unable to read this document.  It may be corrupt.");
                
            }

        }

       
        
    }
}
