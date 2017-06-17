namespace DataScraper
{
    partial class Form1
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
            this.clickMe = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // clickMe
            // 
            this.clickMe.BackColor = System.Drawing.SystemColors.MenuHighlight;
            this.clickMe.Location = new System.Drawing.Point(62, 139);
            this.clickMe.Name = "clickMe";
            this.clickMe.Size = new System.Drawing.Size(76, 37);
            this.clickMe.TabIndex = 1;
            this.clickMe.Text = "Click Me";
            this.clickMe.UseVisualStyleBackColor = false;
            this.clickMe.Click += new System.EventHandler(this.clickMe_Click);
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 261);
            this.Controls.Add(this.clickMe);
            this.Name = "Form1";
            this.Text = "Form1";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button clickMe;
    }
}

