namespace SalesAndMarketingUtilsForm
{
    partial class SaveAsForm
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
            this.btnSaveAs = new System.Windows.Forms.Button();
            this.textBoxFileName = new System.Windows.Forms.TextBox();
            this.SuspendLayout();
            // 
            // btnSaveAs
            // 
            this.btnSaveAs.Location = new System.Drawing.Point(378, 114);
            this.btnSaveAs.Name = "btnSaveAs";
            this.btnSaveAs.Size = new System.Drawing.Size(136, 38);
            this.btnSaveAs.TabIndex = 0;
            this.btnSaveAs.Text = "Save As...";
            this.btnSaveAs.UseVisualStyleBackColor = true;
            this.btnSaveAs.Click += new System.EventHandler(this.Button1_Click);
            // 
            // textBoxFileName
            // 
            this.textBoxFileName.Enabled = false;
            this.textBoxFileName.Location = new System.Drawing.Point(12, 25);
            this.textBoxFileName.Name = "textBoxFileName";
            this.textBoxFileName.Size = new System.Drawing.Size(492, 22);
            this.textBoxFileName.TabIndex = 1;
            this.textBoxFileName.TextChanged += new System.EventHandler(this.textBoxFileName_TextChanged);
            // 
            // SaveAsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(526, 164);
            this.Controls.Add(this.textBoxFileName);
            this.Controls.Add(this.btnSaveAs);
            this.Name = "SaveAsForm";
            this.Text = "Save File";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button btnSaveAs;
        private System.Windows.Forms.TextBox textBoxFileName;
    }
}

