
//Displays Nouns and Noun Phrases of Documents with their tag

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace opening_word_document
{
    public partial class POSTagged : Form
    {
        public POSTagged()
        {
            InitializeComponent();
        }

        public string PosTaggedText { get; set; }

        private void POSTagged_Load(object sender, EventArgs e)
        {
            POSTextbox.Text = PosTaggedText;
        }
    }
}
