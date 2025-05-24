namespace ExcelWordImageUploader
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

        #region Windows Form Designer ChatGPT generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.txtRepo = new System.Windows.Forms.TextBox();
            this.txtExcel = new System.Windows.Forms.TextBox();
            this.txtWord = new System.Windows.Forms.TextBox();
            this.txtSheet = new System.Windows.Forms.TextBox();
            this.btnRun = new System.Windows.Forms.Button();
            this.lblRepo = new System.Windows.Forms.Label();
            this.lblExcel = new System.Windows.Forms.Label();
            this.lblWord = new System.Windows.Forms.Label();
            this.lblSheet = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // txtRepo
            // 
            this.txtRepo.Location = new System.Drawing.Point(160, 15);
            this.txtRepo.Name = "txtRepo";
            this.txtRepo.Size = new System.Drawing.Size(450, 22);
            this.txtRepo.TabIndex = 0;
            this.txtRepo.Text = "Gradescan/images";
            // 
            // txtExcel
            // 
            this.txtExcel.Location = new System.Drawing.Point(160, 45);
            this.txtExcel.Name = "txtExcel";
            this.txtExcel.Size = new System.Drawing.Size(450, 22);
            this.txtExcel.TabIndex = 1;
            this.txtExcel.Text = @"C:\Path\To\File.xlsx";
            // 
            // txtWord
            // 
            this.txtWord.Location = new System.Drawing.Point(160, 75);
            this.txtWord.Name = "txtWord";
            this.txtWord.Size = new System.Drawing.Size(450, 22);
            this.txtWord.TabIndex = 2;
            this.txtWord.Text = @"C:\Path\To\File.docx";
            // 
            // txtSheet
            // 
            this.txtSheet.Location = new System.Drawing.Point(160, 105);
            this.txtSheet.Name = "txtSheet";
            this.txtSheet.Size = new System.Drawing.Size(450, 22);
            this.txtSheet.TabIndex = 3;
            this.txtSheet.Text = "Sheet1";
            // 
            // btnRun
            // 
            this.btnRun.Location = new System.Drawing.Point(160, 140);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(450, 30);
            this.btnRun.TabIndex = 4;
            this.btnRun.Text = "Start Upload";
            this.btnRun.UseVisualStyleBackColor = true;
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // lblRepo
            // 
            this.lblRepo.AutoSize = true;
            this.lblRepo.Location = new System.Drawing.Point(12, 18);
            this.lblRepo.Name = "lblRepo";
            this.lblRepo.Size = new System.Drawing.Size(119, 17);
            this.lblRepo.TabIndex = 5;
            this.lblRepo.Text = "GitHub Repository:";
            // 
            // lblExcel
            // 
            this.lblExcel.AutoSize = true;
            this.lblExcel.Location = new System.Drawing.Point(12, 48);
            this.lblExcel.Name = "lblExcel";
            this.lblExcel.Size = new System.Drawing.Size(116, 17);
            this.lblExcel.TabIndex = 6;
            this.lblExcel.Text = "Excel File (.xlsx):";
            // 
            // lblWord
            // 
            this.lblWord.AutoSize = true;
            this.lblWord.Location = new System.Drawing.Point(12, 78);
            this.lblWord.Name = "lblWord";
            this.lblWord.Size = new System.Drawing.Size(113, 17);
            this.lblWord.TabIndex = 7;
            this.lblWord.Text = "Word File (.docx):";
            // 
            // lblSheet
            // 
            this.lblSheet.AutoSize = true;
            this.lblSheet.Location = new System.Drawing.Point(12, 108);
            this.lblSheet.Name = "lblSheet";
            this.lblSheet.Size = new System.Drawing.Size(113, 17);
            this.lblSheet.TabIndex = 8;
            this.lblSheet.Text = "Worksheet Name:";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(630, 190);
            this.Controls.Add(this.lblSheet);
            this.Controls.Add(this.lblWord);
            this.Controls.Add(this.lblExcel);
            this.Controls.Add(this.lblRepo);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.txtSheet);
            this.Controls.Add(this.txtWord);
            this.Controls.Add(this.txtExcel);
            this.Controls.Add(this.txtRepo);
            this.Name = "Form1";
            this.Text = "Excel to GitHub Uploader";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        #endregion

        private System.Windows.Forms.TextBox txtRepo;
        private System.Windows.Forms.TextBox txtExcel;
        private System.Windows.Forms.TextBox txtWord;
        private System.Windows.Forms.TextBox txtSheet;
        private System.Windows.Forms.Button btnRun;
        private System.Windows.Forms.Label lblRepo;
        private System.Windows.Forms.Label lblExcel;
        private System.Windows.Forms.Label lblWord;
        private System.Windows.Forms.Label lblSheet;
    }
}
