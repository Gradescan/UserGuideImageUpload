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
            this.txtExcelApp = new System.Windows.Forms.TextBox();
            this.txtWordApp = new System.Windows.Forms.TextBox();
            this.btnUploadImages = new System.Windows.Forms.Button();
            this.lblExcel = new System.Windows.Forms.Label();
            this.lblWord = new System.Windows.Forms.Label();
            this.lblSheet = new System.Windows.Forms.Label();
            this.btnBrowseExcel = new System.Windows.Forms.Button();
            this.btnBrowseWord = new System.Windows.Forms.Button();
            this.picBoxGitRepoImage = new System.Windows.Forms.PictureBox();
            this.labelExistingImage = new System.Windows.Forms.Label();
            this.labelWordImage = new System.Windows.Forms.Label();
            this.picBoxWordImage = new System.Windows.Forms.PictureBox();
            this.picBoxPanel = new System.Windows.Forms.Panel();
            this.listBoxCollisions = new System.Windows.Forms.ListBox();
            this.labelCollisions = new System.Windows.Forms.Label();
            this.comboBoxWorksheet = new System.Windows.Forms.ComboBox();
            this.buttonStop = new System.Windows.Forms.Button();
            this.labelFileName = new System.Windows.Forms.Label();
            this.btnClearAltText = new System.Windows.Forms.Button();
            this.btnVerify = new System.Windows.Forms.Button();
            this.btnAssign = new System.Windows.Forms.Button();
            this.lblStatus = new System.Windows.Forms.Label();
            this.labelStatus = new System.Windows.Forms.Label();
            this.btnCreateAIFile = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.picBoxGitRepoImage)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picBoxWordImage)).BeginInit();
            this.picBoxPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtExcelApp
            // 
            this.txtExcelApp.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.txtExcelApp.Location = new System.Drawing.Point(170, 106);
            this.txtExcelApp.Name = "txtExcelApp";
            this.txtExcelApp.Size = new System.Drawing.Size(959, 29);
            this.txtExcelApp.TabIndex = 5;
            this.txtExcelApp.Text = "C:\\Users\\Tim\\Documents\\__ngTTMv831\\ngTTM\\angular\\src\\assets\\docs\\User Guides Imag" +
    "e Map.xlsx";
            // 
            // txtWordApp
            // 
            this.txtWordApp.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.txtWordApp.Location = new System.Drawing.Point(170, 16);
            this.txtWordApp.Name = "txtWordApp";
            this.txtWordApp.Size = new System.Drawing.Size(959, 29);
            this.txtWordApp.TabIndex = 6;
            // 
            // btnUploadImages
            // 
            this.btnUploadImages.BackColor = System.Drawing.Color.LawnGreen;
            this.btnUploadImages.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnUploadImages.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUploadImages.Location = new System.Drawing.Point(653, 144);
            this.btnUploadImages.Name = "btnUploadImages";
            this.btnUploadImages.Size = new System.Drawing.Size(168, 59);
            this.btnUploadImages.TabIndex = 8;
            this.btnUploadImages.Text = "Upload Images";
            this.btnUploadImages.UseVisualStyleBackColor = false;
            this.btnUploadImages.Click += new System.EventHandler(this.btnUploadImages_Click);
            // 
            // lblExcel
            // 
            this.lblExcel.AutoSize = true;
            this.lblExcel.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.lblExcel.Location = new System.Drawing.Point(28, 109);
            this.lblExcel.Name = "lblExcel";
            this.lblExcel.Size = new System.Drawing.Size(117, 21);
            this.lblExcel.TabIndex = 1;
            this.lblExcel.Text = "Excel File (.xlsx):";
            // 
            // lblWord
            // 
            this.lblWord.AutoSize = true;
            this.lblWord.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.lblWord.Location = new System.Drawing.Point(17, 19);
            this.lblWord.Name = "lblWord";
            this.lblWord.Size = new System.Drawing.Size(128, 21);
            this.lblWord.TabIndex = 2;
            this.lblWord.Text = "Word File (.docx):";
            // 
            // lblSheet
            // 
            this.lblSheet.AutoSize = true;
            this.lblSheet.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.lblSheet.Location = new System.Drawing.Point(12, 144);
            this.lblSheet.Name = "lblSheet";
            this.lblSheet.Size = new System.Drawing.Size(133, 21);
            this.lblSheet.TabIndex = 3;
            this.lblSheet.Text = "Worksheet Name:";
            // 
            // btnBrowseExcel
            // 
            this.btnBrowseExcel.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.btnBrowseExcel.Location = new System.Drawing.Point(1135, 105);
            this.btnBrowseExcel.Name = "btnBrowseExcel";
            this.btnBrowseExcel.Size = new System.Drawing.Size(40, 29);
            this.btnBrowseExcel.TabIndex = 10;
            this.btnBrowseExcel.Text = "...";
            this.btnBrowseExcel.Click += new System.EventHandler(this.btnBrowseExcel_Click);
            // 
            // btnBrowseWord
            // 
            this.btnBrowseWord.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.btnBrowseWord.Location = new System.Drawing.Point(1135, 15);
            this.btnBrowseWord.Name = "btnBrowseWord";
            this.btnBrowseWord.Size = new System.Drawing.Size(40, 29);
            this.btnBrowseWord.TabIndex = 11;
            this.btnBrowseWord.Text = "...";
            this.btnBrowseWord.Click += new System.EventHandler(this.btnBrowseWord_Click);
            // 
            // picBoxGitRepoImage
            // 
            this.picBoxGitRepoImage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.picBoxGitRepoImage.Location = new System.Drawing.Point(8, 7);
            this.picBoxGitRepoImage.Name = "picBoxGitRepoImage";
            this.picBoxGitRepoImage.Size = new System.Drawing.Size(579, 350);
            this.picBoxGitRepoImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picBoxGitRepoImage.TabIndex = 12;
            this.picBoxGitRepoImage.TabStop = false;
            // 
            // labelExistingImage
            // 
            this.labelExistingImage.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.labelExistingImage.AutoSize = true;
            this.labelExistingImage.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.labelExistingImage.Location = new System.Drawing.Point(205, 295);
            this.labelExistingImage.Name = "labelExistingImage";
            this.labelExistingImage.Size = new System.Drawing.Size(172, 21);
            this.labelExistingImage.TabIndex = 13;
            this.labelExistingImage.Text = "Existing git Repo Image";
            this.labelExistingImage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // labelWordImage
            // 
            this.labelWordImage.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelWordImage.AutoSize = true;
            this.labelWordImage.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.labelWordImage.Location = new System.Drawing.Point(836, 291);
            this.labelWordImage.Name = "labelWordImage";
            this.labelWordImage.Size = new System.Drawing.Size(95, 21);
            this.labelWordImage.TabIndex = 14;
            this.labelWordImage.Text = "Word Image";
            this.labelWordImage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // picBoxWordImage
            // 
            this.picBoxWordImage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.picBoxWordImage.Location = new System.Drawing.Point(610, 7);
            this.picBoxWordImage.Name = "picBoxWordImage";
            this.picBoxWordImage.Size = new System.Drawing.Size(565, 350);
            this.picBoxWordImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picBoxWordImage.TabIndex = 15;
            this.picBoxWordImage.TabStop = false;
            // 
            // picBoxPanel
            // 
            this.picBoxPanel.BackColor = System.Drawing.SystemColors.Control;
            this.picBoxPanel.Controls.Add(this.picBoxWordImage);
            this.picBoxPanel.Controls.Add(this.picBoxGitRepoImage);
            this.picBoxPanel.Location = new System.Drawing.Point(3, 323);
            this.picBoxPanel.Name = "picBoxPanel";
            this.picBoxPanel.Size = new System.Drawing.Size(1179, 365);
            this.picBoxPanel.TabIndex = 16;
            // 
            // listBoxCollisions
            // 
            this.listBoxCollisions.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.listBoxCollisions.FormattingEnabled = true;
            this.listBoxCollisions.ItemHeight = 21;
            this.listBoxCollisions.Location = new System.Drawing.Point(170, 184);
            this.listBoxCollisions.Name = "listBoxCollisions";
            this.listBoxCollisions.Size = new System.Drawing.Size(385, 88);
            this.listBoxCollisions.TabIndex = 17;
            // 
            // labelCollisions
            // 
            this.labelCollisions.AutoSize = true;
            this.labelCollisions.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelCollisions.Location = new System.Drawing.Point(65, 203);
            this.labelCollisions.Name = "labelCollisions";
            this.labelCollisions.Size = new System.Drawing.Size(80, 21);
            this.labelCollisions.TabIndex = 18;
            this.labelCollisions.Text = "Collisions:";
            // 
            // comboBoxWorksheet
            // 
            this.comboBoxWorksheet.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBoxWorksheet.FormattingEnabled = true;
            this.comboBoxWorksheet.Location = new System.Drawing.Point(170, 147);
            this.comboBoxWorksheet.Name = "comboBoxWorksheet";
            this.comboBoxWorksheet.Size = new System.Drawing.Size(385, 29);
            this.comboBoxWorksheet.TabIndex = 19;
            // 
            // buttonStop
            // 
            this.buttonStop.BackColor = System.Drawing.Color.Red;
            this.buttonStop.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonStop.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonStop.Location = new System.Drawing.Point(986, 144);
            this.buttonStop.Name = "buttonStop";
            this.buttonStop.Size = new System.Drawing.Size(143, 104);
            this.buttonStop.TabIndex = 20;
            this.buttonStop.Text = "STOP";
            this.buttonStop.UseVisualStyleBackColor = false;
            this.buttonStop.Click += new System.EventHandler(this.buttonStop_Click);
            // 
            // labelFileName
            // 
            this.labelFileName.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelFileName.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Bold);
            this.labelFileName.Location = new System.Drawing.Point(399, 291);
            this.labelFileName.Name = "labelFileName";
            this.labelFileName.Size = new System.Drawing.Size(392, 28);
            this.labelFileName.TabIndex = 21;
            this.labelFileName.Text = "file name";
            this.labelFileName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnClearAltText
            // 
            this.btnClearAltText.BackColor = System.Drawing.Color.Yellow;
            this.btnClearAltText.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnClearAltText.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClearAltText.Location = new System.Drawing.Point(170, 51);
            this.btnClearAltText.Name = "btnClearAltText";
            this.btnClearAltText.Size = new System.Drawing.Size(223, 46);
            this.btnClearAltText.TabIndex = 22;
            this.btnClearAltText.Text = "Clear Alt Text";
            this.btnClearAltText.UseVisualStyleBackColor = false;
            this.btnClearAltText.Click += new System.EventHandler(this.btnClearAltText_Click);
            // 
            // btnVerify
            // 
            this.btnVerify.BackColor = System.Drawing.Color.Aqua;
            this.btnVerify.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnVerify.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnVerify.Location = new System.Drawing.Point(903, 51);
            this.btnVerify.Name = "btnVerify";
            this.btnVerify.Size = new System.Drawing.Size(226, 46);
            this.btnVerify.TabIndex = 23;
            this.btnVerify.Text = "Verify Alt Text";
            this.btnVerify.UseVisualStyleBackColor = false;
            this.btnVerify.Click += new System.EventHandler(this.btnVerify_Click);
            // 
            // btnAssign
            // 
            this.btnAssign.BackColor = System.Drawing.Color.Fuchsia;
            this.btnAssign.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnAssign.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAssign.Location = new System.Drawing.Point(484, 51);
            this.btnAssign.Name = "btnAssign";
            this.btnAssign.Size = new System.Drawing.Size(225, 46);
            this.btnAssign.TabIndex = 24;
            this.btnAssign.Text = "Assign Alt Text";
            this.btnAssign.UseVisualStyleBackColor = false;
            this.btnAssign.Click += new System.EventHandler(this.btnAssign_Click);
            // 
            // lblStatus
            // 
            this.lblStatus.AutoSize = true;
            this.lblStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.lblStatus.Location = new System.Drawing.Point(838, 167);
            this.lblStatus.Name = "lblStatus";
            this.lblStatus.Size = new System.Drawing.Size(60, 24);
            this.lblStatus.TabIndex = 25;
            this.lblStatus.Text = "Status";
            // 
            // labelStatus
            // 
            this.labelStatus.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelStatus.AutoSize = true;
            this.labelStatus.Font = new System.Drawing.Font("Microsoft Sans Serif", 14F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelStatus.Location = new System.Drawing.Point(838, 198);
            this.labelStatus.Name = "labelStatus";
            this.labelStatus.Size = new System.Drawing.Size(25, 24);
            this.labelStatus.TabIndex = 27;
            this.labelStatus.Text = "...";
            // 
            // btnCreateAIFile
            // 
            this.btnCreateAIFile.BackColor = System.Drawing.Color.Lime;
            this.btnCreateAIFile.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnCreateAIFile.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnCreateAIFile.Location = new System.Drawing.Point(653, 218);
            this.btnCreateAIFile.Name = "btnCreateAIFile";
            this.btnCreateAIFile.Size = new System.Drawing.Size(168, 59);
            this.btnCreateAIFile.TabIndex = 28;
            this.btnCreateAIFile.Text = "Create .txt File";
            this.btnCreateAIFile.UseVisualStyleBackColor = false;
            this.btnCreateAIFile.Click += new System.EventHandler(this.btnCreateTxtFile_Click);
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(1184, 691);
            this.Controls.Add(this.btnCreateAIFile);
            this.Controls.Add(this.labelStatus);
            this.Controls.Add(this.lblStatus);
            this.Controls.Add(this.btnAssign);
            this.Controls.Add(this.btnVerify);
            this.Controls.Add(this.btnClearAltText);
            this.Controls.Add(this.labelFileName);
            this.Controls.Add(this.buttonStop);
            this.Controls.Add(this.comboBoxWorksheet);
            this.Controls.Add(this.labelCollisions);
            this.Controls.Add(this.listBoxCollisions);
            this.Controls.Add(this.picBoxPanel);
            this.Controls.Add(this.labelWordImage);
            this.Controls.Add(this.labelExistingImage);
            this.Controls.Add(this.lblExcel);
            this.Controls.Add(this.lblWord);
            this.Controls.Add(this.lblSheet);
            this.Controls.Add(this.txtExcelApp);
            this.Controls.Add(this.txtWordApp);
            this.Controls.Add(this.btnUploadImages);
            this.Controls.Add(this.btnBrowseExcel);
            this.Controls.Add(this.btnBrowseWord);
            this.Name = "Form1";
            this.Text = "Excel to GitHub Uploader";
            ((System.ComponentModel.ISupportInitialize)(this.picBoxGitRepoImage)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picBoxWordImage)).EndInit();
            this.picBoxPanel.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox txtExcelApp;
        private System.Windows.Forms.TextBox txtWordApp;
        private System.Windows.Forms.Button btnUploadImages;
        private System.Windows.Forms.Label lblExcel;
        private System.Windows.Forms.Label lblWord;
        private System.Windows.Forms.Label lblSheet;
        private System.Windows.Forms.Button btnBrowseExcel;
        private System.Windows.Forms.Button btnBrowseWord;
        private System.Windows.Forms.PictureBox picBoxGitRepoImage;
        private System.Windows.Forms.Label labelExistingImage;
        private System.Windows.Forms.Label labelWordImage;
        private System.Windows.Forms.PictureBox picBoxWordImage;
        private System.Windows.Forms.Panel picBoxPanel;
        private System.Windows.Forms.ListBox listBoxCollisions;
        private System.Windows.Forms.Label labelCollisions;
        private System.Windows.Forms.ComboBox comboBoxWorksheet;
        private System.Windows.Forms.Button buttonStop;
        private System.Windows.Forms.Label labelFileName;
        private System.Windows.Forms.Button btnClearAltText;
        private System.Windows.Forms.Button btnVerify;
        private System.Windows.Forms.Button btnAssign;
        private System.Windows.Forms.Label lblStatus;
        private System.Windows.Forms.Label labelStatus;
        private System.Windows.Forms.Button btnCreateAIFile;
    }
}
