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

using System.Runtime.InteropServices;

namespace Word2Pdf
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void BrowseBtn_Click(object sender, EventArgs e)
        {

             OpenFileDialog ofd=new OpenFileDialog();
             ofd.Filter = "Word Document(*.doc) | *.doc";

             if (ofd.ShowDialog() == DialogResult.OK)
             {

                 filePath.Text = ofd.FileName;
                 ApplicationClass wordApplication = new ApplicationClass();

                 Document wordDocument = null;
                 object paramSourceDocPath = filePath.Text;
                 object paramMissing = Type.Missing;
                 string Direct = Path.GetDirectoryName(filePath.Text);
                 string fwext = Path.GetFileNameWithoutExtension(filePath.Text);
                 string pwext = Direct + fwext;

                // MessageBox.Show(pwext);

                 string paramExportFilePath = pwext + ".pdf";
                 WdExportFormat paramExportFormat = WdExportFormat.wdExportFormatPDF;
                 bool paramOpenAfterExport = false;
                 WdExportOptimizeFor paramExportOptimizeFor =
                     WdExportOptimizeFor.wdExportOptimizeForPrint;
                 WdExportRange paramExportRange = WdExportRange.wdExportAllDocument;
                 int paramStartPage = 0;
                 int paramEndPage = 0;
                 WdExportItem paramExportItem = WdExportItem.wdExportDocumentContent;
                 bool paramIncludeDocProps = true;
                 bool paramKeepIRM = true;
                 WdExportCreateBookmarks paramCreateBookmarks =
                     WdExportCreateBookmarks.wdExportCreateWordBookmarks;
                 bool paramDocStructureTags = true;
                 bool paramBitmapMissingFonts = true;
                 bool paramUseISO19005_1 = false;
                 try
                 {
                     // Open the source document.
                     wordDocument = wordApplication.Documents.Open(
                         ref paramSourceDocPath, ref paramMissing, ref paramMissing,
                         ref paramMissing, ref paramMissing, ref paramMissing,
                         ref paramMissing, ref paramMissing, ref paramMissing,
                         ref paramMissing, ref paramMissing, ref paramMissing,
                         ref paramMissing, ref paramMissing, ref paramMissing,
                         ref paramMissing);

                     // Export it in the specified format.
                     if (wordDocument != null)
                         wordDocument.ExportAsFixedFormat(paramExportFilePath,
                                                          paramExportFormat, paramOpenAfterExport,
                                                          paramExportOptimizeFor, paramExportRange, paramStartPage,
                                                          paramEndPage, paramExportItem, paramIncludeDocProps,
                                                          paramKeepIRM, paramCreateBookmarks, paramDocStructureTags,
                                                          paramBitmapMissingFonts, paramUseISO19005_1,
                                                          ref paramMissing);
                 }
                 catch (Exception ex)
                 {
                     // Respond to the error
                 }
                 finally
                 {
                     // Close and release the Document object.
                     if (wordDocument != null)
                     {
                         wordDocument.Close(ref paramMissing, ref paramMissing,
                                            ref paramMissing);
                         wordDocument = null;
                     }

                     // Quit Word and release the ApplicationClass object.
                     if (wordApplication != null)
                     {
                         wordApplication.Quit(ref paramMissing, ref paramMissing,
                                              ref paramMissing);
                         wordApplication = null;
                     }

                     GC.Collect();
                     GC.WaitForPendingFinalizers();
                     GC.Collect();
                     GC.WaitForPendingFinalizers();
                 }
             }
        }
    }
}
