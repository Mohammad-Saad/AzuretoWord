namespace HtmlToWord
{
    partial class AzureConverter
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
            this.button1 = new System.Windows.Forms.Button();
            this.txthtmlPath = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.txtScreenShots = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(358, 177);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 0;
            this.button1.Text = "Convert";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // txthtmlPath
            // 
            this.txthtmlPath.Location = new System.Drawing.Point(155, 86);
            this.txthtmlPath.Name = "txthtmlPath";
            this.txthtmlPath.Size = new System.Drawing.Size(278, 20);
            this.txthtmlPath.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(44, 92);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Html Path";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // txtScreenShots
            // 
            this.txtScreenShots.Location = new System.Drawing.Point(155, 127);
            this.txtScreenShots.Name = "txtScreenShots";
            this.txtScreenShots.Size = new System.Drawing.Size(277, 20);
            this.txtScreenShots.TabIndex = 3;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(47, 127);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(93, 13);
            this.label2.TabIndex = 4;
            this.label2.Text = "ScreenShots Path";
            // 
            // AzureConverter
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(473, 320);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtScreenShots);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.txthtmlPath);
            this.Controls.Add(this.button1);
            this.Name = "AzureConverter";
            this.Text = "AzureConverter";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.TextBox txthtmlPath;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox txtScreenShots;
        private System.Windows.Forms.Label label2;
    }
}

