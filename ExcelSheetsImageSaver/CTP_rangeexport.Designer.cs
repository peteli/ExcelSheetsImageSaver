namespace ExcelSheetsImageSaver
{
    partial class CTP_rangeexport
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
            this.TheLabel = new System.Windows.Forms.Label();
            this.checkedListBox1 = new System.Windows.Forms.CheckedListBox();
            this.SuspendLayout();
            // 
            // TheLabel
            // 
            this.TheLabel.AutoSize = true;
            this.TheLabel.Location = new System.Drawing.Point(13, 13);
            this.TheLabel.Name = "TheLabel";
            this.TheLabel.Size = new System.Drawing.Size(35, 13);
            this.TheLabel.TabIndex = 0;
            this.TheLabel.Text = "label1";
            // 
            // checkedListBox1
            // 
            this.checkedListBox1.FormattingEnabled = true;
            this.checkedListBox1.Location = new System.Drawing.Point(12, 45);
            this.checkedListBox1.Name = "checkedListBox1";
            this.checkedListBox1.Size = new System.Drawing.Size(260, 124);
            this.checkedListBox1.TabIndex = 1;
            // 
            // CTP_rangeexport
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.checkedListBox1);
            this.Controls.Add(this.TheLabel);
            this.Name = "CTP_rangeexport";
            this.Text = "CTP_rangeexport";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label TheLabel;
        private System.Windows.Forms.CheckedListBox checkedListBox1;
    }
}