namespace ExcelWordImageUploader
{
    partial class Form1
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null)) components.Dispose();
            base.Dispose(disposing);
        }

        #region Windows Form Designer ChatGPT generated code

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
            this.btnBrowseExcel = new System.Windows.Forms.Button();
            this.btnBrowseWord = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // txtRepo
            // 
            this.txtRepo.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.txtRepo.Location = new System.Drawing.Point(220, 15);
            this.txtRepo.Name = "txtRepo";
            this.txtRepo.Size = new System.Drawing.Size(876, 29);
            this.txtRepo.TabIndex = 4;
            this.txtRepo.Text = "Professional Edition";
            // 
            // txtExcel
            // 
            this.txtExcel.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.txtExcel.Location = new System.Drawing.Point(220, 115);
            this.txtExcel.Name = "txtExcel";
            this.txtExcel.Size = new System.Drawing.Size(876, 29);
            this.txtExcel.TabIndex = 5;
            this.txtExcel.Text = "C:\\Users\\Tim\\Documents\\__ngTTMv831\\ngTTM\\angular\\src\\assets\\docs\\Professional Use" +
    "r Guide";
            // 
            // txtWord
            // 
            this.txtWord.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.txtWord.Location = new System.Drawing.Point(220, 64);
            this.txtWord.Name = "txtWord";
            this.txtWord.Size = new System.Drawing.Size(876, 29);
            this.txtWord.TabIndex = 6;
            this.txtWord.Text = "C:\\Users\\Tim\\Documents\\__ngTTMv831\\ngTTM\\angular\\src\\assets\\docs\\Professional Use" +
    "r Guide";
            // 
            // txtSheet
            // 
            this.txtSheet.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.txtSheet.Location = new System.Drawing.Point(220, 167);
            this.txtSheet.Name = "txtSheet";
            this.txtSheet.Size = new System.Drawing.Size(708, 29);
            this.txtSheet.TabIndex = 7;
            this.txtSheet.Text = "Professional Edition";
            // 
            // btnRun
            // 
            this.btnRun.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.btnRun.Location = new System.Drawing.Point(496, 218);
            this.btnRun.Name = "btnRun";
            this.btnRun.Size = new System.Drawing.Size(210, 66);
            this.btnRun.TabIndex = 8;
            this.btnRun.Text = "Start Upload";
            this.btnRun.Click += new System.EventHandler(this.btnRun_Click);
            // 
            // lblRepo
            // 
            this.lblRepo.AutoSize = true;
            this.lblRepo.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.lblRepo.Location = new System.Drawing.Point(12, 18);
            this.lblRepo.Name = "lblRepo";
            this.lblRepo.Size = new System.Drawing.Size(154, 21);
            this.lblRepo.TabIndex = 0;
            this.lblRepo.Text = "GitHub media folder:";
            // 
            // lblExcel
            // 
            this.lblExcel.AutoSize = true;
            this.lblExcel.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.lblExcel.Location = new System.Drawing.Point(12, 118);
            this.lblExcel.Name = "lblExcel";
            this.lblExcel.Size = new System.Drawing.Size(117, 21);
            this.lblExcel.TabIndex = 1;
            this.lblExcel.Text = "Excel File (.xlsx):";
            // 
            // lblWord
            // 
            this.lblWord.AutoSize = true;
            this.lblWord.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.lblWord.Location = new System.Drawing.Point(12, 67);
            this.lblWord.Name = "lblWord";
            this.lblWord.Size = new System.Drawing.Size(128, 21);
            this.lblWord.TabIndex = 2;
            this.lblWord.Text = "Word File (.docx):";
            // 
            // lblSheet
            // 
            this.lblSheet.AutoSize = true;
            this.lblSheet.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.lblSheet.Location = new System.Drawing.Point(12, 170);
            this.lblSheet.Name = "lblSheet";
            this.lblSheet.Size = new System.Drawing.Size(133, 21);
            this.lblSheet.TabIndex = 3;
            this.lblSheet.Text = "Worksheet Name:";
            // 
            // btnBrowseExcel
            // 
            this.btnBrowseExcel.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.btnBrowseExcel.Location = new System.Drawing.Point(1128, 115);
            this.btnBrowseExcel.Name = "btnBrowseExcel";
            this.btnBrowseExcel.Size = new System.Drawing.Size(40, 29);
            this.btnBrowseExcel.TabIndex = 10;
            this.btnBrowseExcel.Text = "...";
            this.btnBrowseExcel.Click += new System.EventHandler(this.btnBrowseExcel_Click);
            // 
            // btnBrowseWord
            // 
            this.btnBrowseWord.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.btnBrowseWord.Location = new System.Drawing.Point(1128, 64);
            this.btnBrowseWord.Name = "btnBrowseWord";
            this.btnBrowseWord.Size = new System.Drawing.Size(40, 29);
            this.btnBrowseWord.TabIndex = 11;
            this.btnBrowseWord.Text = "...";
            this.btnBrowseWord.Click += new System.EventHandler(this.btnBrowseWord_Click);
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(1196, 307);
            this.Controls.Add(this.lblRepo);
            this.Controls.Add(this.lblExcel);
            this.Controls.Add(this.lblWord);
            this.Controls.Add(this.lblSheet);
            this.Controls.Add(this.txtRepo);
            this.Controls.Add(this.txtExcel);
            this.Controls.Add(this.txtWord);
            this.Controls.Add(this.txtSheet);
            this.Controls.Add(this.btnRun);
            this.Controls.Add(this.btnBrowseExcel);
            this.Controls.Add(this.btnBrowseWord);
            this.Name = "Form1";
            this.Text = "Excel to GitHub Uploader";
            this.Load += new System.EventHandler(this.Form1_Load_1);
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
        private System.Windows.Forms.Button btnBrowseExcel;
        private System.Windows.Forms.Button btnBrowseWord;
    }
}
