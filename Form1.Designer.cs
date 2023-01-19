namespace Esker_Inv_Integration
{
    partial class FTS00OII
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.Log = new System.Windows.Forms.RichTextBox();
            this.btnExecute = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // Log
            // 
            this.Log.Location = new System.Drawing.Point(24, 22);
            this.Log.Name = "Log";
            this.Log.Size = new System.Drawing.Size(729, 151);
            this.Log.TabIndex = 0;
            this.Log.Text = "";
            // 
            // btnExecute
            // 
            this.btnExecute.Location = new System.Drawing.Point(24, 179);
            this.btnExecute.Name = "btnExecute";
            this.btnExecute.Size = new System.Drawing.Size(729, 23);
            this.btnExecute.TabIndex = 1;
            this.btnExecute.Text = "Execute";
            this.btnExecute.UseVisualStyleBackColor = true;
            this.btnExecute.Click += new System.EventHandler(this.btnExecute_Click);
            // 
            // FTS00OII
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(769, 226);
            this.Controls.Add(this.btnExecute);
            this.Controls.Add(this.Log);
            this.Name = "FTS00OII";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Esker Invoice Integration";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FTS00OII_FormClosing);
            this.Load += new System.EventHandler(this.FTS00OII_Load);
            this.ResumeLayout(false);

        }

        #endregion

        public System.Windows.Forms.RichTextBox Log;
        private System.Windows.Forms.Button btnExecute;
    }
}

