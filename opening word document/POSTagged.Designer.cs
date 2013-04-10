namespace opening_word_document
{
    partial class POSTagged
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
            this.POSTextbox = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // POSTextbox
            // 
            this.POSTextbox.Location = new System.Drawing.Point(12, 12);
            this.POSTextbox.Name = "POSTextbox";
            this.POSTextbox.Size = new System.Drawing.Size(504, 421);
            this.POSTextbox.TabIndex = 0;
            this.POSTextbox.Text = "";
            // 
            // POSTagged
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(528, 445);
            this.Controls.Add(this.POSTextbox);
            this.Name = "POSTagged";
            this.Text = "POSTagged";
            this.Load += new System.EventHandler(this.POSTagged_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.RichTextBox POSTextbox;
    }
}