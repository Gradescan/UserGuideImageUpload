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
            this.btnUpload = new System.Windows.Forms.Button();
            this.lblExcel = new System.Windows.Forms.Label();
            this.lblWord = new System.Windows.Forms.Label();
            this.lblSheet = new System.Windows.Forms.Label();
            this.btnBrowseExcel = new System.Windows.Forms.Button();
            this.btnBrowseWord = new System.Windows.Forms.Button();
            this.picBoxOldImage = new System.Windows.Forms.PictureBox();
            this.labelOldImage = new System.Windows.Forms.Label();
            this.labelNewImage = new System.Windows.Forms.Label();
            this.picBoxNewImage = new System.Windows.Forms.PictureBox();
            this.picBoxPanel = new System.Windows.Forms.Panel();
            this.listBoxCollisions = new System.Windows.Forms.ListBox();
            this.labelCollisions = new System.Windows.Forms.Label();
            this.comboBoxWorksheetNames = new System.Windows.Forms.ComboBox();
            this.buttonStop = new System.Windows.Forms.Button();
            this.labelFileName = new System.Windows.Forms.Label();
            this.btnClearAltText = new System.Windows.Forms.Button();
            this.btnVerify = new System.Windows.Forms.Button();
            this.btnAssign = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.picBoxOldImage)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.picBoxNewImage)).BeginInit();
            this.picBoxPanel.SuspendLayout();
            this.SuspendLayout();
            // 
            // txtExcelApp
            // 
            this.txtExcelApp.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.txtExcelApp.Location = new System.Drawing.Point(170, 51);
            this.txtExcelApp.Name = "txtExcelApp";
            this.txtExcelApp.Size = new System.Drawing.Size(820, 29);
            this.txtExcelApp.TabIndex = 5;
            this.txtExcelApp.Text = "C:\\Users\\Tim\\Documents\\__ngTTMv831\\ngTTM\\angular\\src\\assets\\docs\\User Guides Imag" +
    "e Map.xlsx";
            this.txtExcelApp.TextChanged += new System.EventHandler(this.txtExcel_TextChanged);
            // 
            // txtWordApp
            // 
            this.txtWordApp.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.txtWordApp.Location = new System.Drawing.Point(170, 16);
            this.txtWordApp.Name = "txtWordApp";
            this.txtWordApp.Size = new System.Drawing.Size(820, 29);
            this.txtWordApp.TabIndex = 6;
            this.txtWordApp.Text = "Select a User Guide";
            // 
            // btnUpload
            // 
            this.btnUpload.BackColor = System.Drawing.Color.LawnGreen;
            this.btnUpload.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnUpload.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnUpload.Location = new System.Drawing.Point(553, 161);
            this.btnUpload.Name = "btnUpload";
            this.btnUpload.Size = new System.Drawing.Size(223, 69);
            this.btnUpload.TabIndex = 8;
            this.btnUpload.Text = "Upload Images";
            this.btnUpload.UseVisualStyleBackColor = false;
            this.btnUpload.Click += new System.EventHandler(this.btnUpload_Click);
            // 
            // lblExcel
            // 
            this.lblExcel.AutoSize = true;
            this.lblExcel.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.lblExcel.Location = new System.Drawing.Point(28, 54);
            this.lblExcel.Name = "lblExcel";
            this.lblExcel.Size = new System.Drawing.Size(117, 21);
            this.lblExcel.TabIndex = 1;
            this.lblExcel.Text = "Excel File (.xlsx):";
            this.lblExcel.Click += new System.EventHandler(this.lblExcel_Click);
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
            this.lblSheet.Location = new System.Drawing.Point(12, 89);
            this.lblSheet.Name = "lblSheet";
            this.lblSheet.Size = new System.Drawing.Size(133, 21);
            this.lblSheet.TabIndex = 3;
            this.lblSheet.Text = "Worksheet Name:";
            // 
            // btnBrowseExcel
            // 
            this.btnBrowseExcel.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.btnBrowseExcel.Location = new System.Drawing.Point(996, 50);
            this.btnBrowseExcel.Name = "btnBrowseExcel";
            this.btnBrowseExcel.Size = new System.Drawing.Size(40, 29);
            this.btnBrowseExcel.TabIndex = 10;
            this.btnBrowseExcel.Text = "...";
            this.btnBrowseExcel.Click += new System.EventHandler(this.btnBrowseExcel_Click);
            // 
            // btnBrowseWord
            // 
            this.btnBrowseWord.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.btnBrowseWord.Location = new System.Drawing.Point(996, 15);
            this.btnBrowseWord.Name = "btnBrowseWord";
            this.btnBrowseWord.Size = new System.Drawing.Size(40, 29);
            this.btnBrowseWord.TabIndex = 11;
            this.btnBrowseWord.Text = "...";
            this.btnBrowseWord.Click += new System.EventHandler(this.btnBrowseWord_Click);
            // 
            // picBoxOldImage
            // 
            this.picBoxOldImage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.picBoxOldImage.Location = new System.Drawing.Point(6, 7);
            this.picBoxOldImage.Name = "picBoxOldImage";
            this.picBoxOldImage.Size = new System.Drawing.Size(476, 383);
            this.picBoxOldImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picBoxOldImage.TabIndex = 12;
            this.picBoxOldImage.TabStop = false;
            // 
            // labelOldImage
            // 
            this.labelOldImage.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.labelOldImage.AutoSize = true;
            this.labelOldImage.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.labelOldImage.Location = new System.Drawing.Point(213, 239);
            this.labelOldImage.Name = "labelOldImage";
            this.labelOldImage.Size = new System.Drawing.Size(82, 21);
            this.labelOldImage.TabIndex = 13;
            this.labelOldImage.Text = "Old Image";
            this.labelOldImage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            this.labelOldImage.Click += new System.EventHandler(this.label1_Click);
            // 
            // labelNewImage
            // 
            this.labelNewImage.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.labelNewImage.AutoSize = true;
            this.labelNewImage.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.labelNewImage.Location = new System.Drawing.Point(737, 239);
            this.labelNewImage.Name = "labelNewImage";
            this.labelNewImage.Size = new System.Drawing.Size(89, 21);
            this.labelNewImage.TabIndex = 14;
            this.labelNewImage.Text = "New Image";
            this.labelNewImage.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // picBoxNewImage
            // 
            this.picBoxNewImage.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.picBoxNewImage.Location = new System.Drawing.Point(541, 7);
            this.picBoxNewImage.Name = "picBoxNewImage";
            this.picBoxNewImage.Size = new System.Drawing.Size(476, 383);
            this.picBoxNewImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.picBoxNewImage.TabIndex = 15;
            this.picBoxNewImage.TabStop = false;
            // 
            // picBoxPanel
            // 
            this.picBoxPanel.BackColor = System.Drawing.SystemColors.Control;
            this.picBoxPanel.Controls.Add(this.picBoxNewImage);
            this.picBoxPanel.Controls.Add(this.picBoxOldImage);
            this.picBoxPanel.Location = new System.Drawing.Point(16, 271);
            this.picBoxPanel.Name = "picBoxPanel";
            this.picBoxPanel.Size = new System.Drawing.Size(1020, 396);
            this.picBoxPanel.TabIndex = 16;
            // 
            // listBoxCollisions
            // 
            this.listBoxCollisions.Font = new System.Drawing.Font("Segoe UI", 12F);
            this.listBoxCollisions.FormattingEnabled = true;
            this.listBoxCollisions.ItemHeight = 21;
            this.listBoxCollisions.Location = new System.Drawing.Point(170, 142);
            this.listBoxCollisions.Name = "listBoxCollisions";
            this.listBoxCollisions.Size = new System.Drawing.Size(328, 88);
            this.listBoxCollisions.TabIndex = 17;
            this.listBoxCollisions.SelectedIndexChanged += new System.EventHandler(this.listBoxCollisions_SelectedIndexChanged);
            // 
            // labelCollisions
            // 
            this.labelCollisions.AutoSize = true;
            this.labelCollisions.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.labelCollisions.Location = new System.Drawing.Point(65, 161);
            this.labelCollisions.Name = "labelCollisions";
            this.labelCollisions.Size = new System.Drawing.Size(80, 21);
            this.labelCollisions.TabIndex = 18;
            this.labelCollisions.Text = "Collisions:";
            // 
            // comboBoxWorksheetNames
            // 
            this.comboBoxWorksheetNames.Font = new System.Drawing.Font("Segoe UI", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.comboBoxWorksheetNames.FormattingEnabled = true;
            this.comboBoxWorksheetNames.Items.AddRange(new object[] {
            "Basic Edition",
            "Professional Edition",
            "Basic Forms",
            "Professional Forms"});
            this.comboBoxWorksheetNames.Location = new System.Drawing.Point(170, 92);
            this.comboBoxWorksheetNames.Name = "comboBoxWorksheetNames";
            this.comboBoxWorksheetNames.Size = new System.Drawing.Size(328, 29);
            this.comboBoxWorksheetNames.TabIndex = 19;
            this.comboBoxWorksheetNames.SelectedIndexChanged += new System.EventHandler(this.comboBoxWorksheetNames_SelectedIndexChanged);
            // 
            // buttonStop
            // 
            this.buttonStop.BackColor = System.Drawing.Color.Red;
            this.buttonStop.Cursor = System.Windows.Forms.Cursors.Hand;
            this.buttonStop.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.buttonStop.Location = new System.Drawing.Point(806, 161);
            this.buttonStop.Name = "buttonStop";
            this.buttonStop.Size = new System.Drawing.Size(223, 69);
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
            this.labelFileName.Location = new System.Drawing.Point(341, 239);
            this.labelFileName.Name = "labelFileName";
            this.labelFileName.Size = new System.Drawing.Size(369, 21);
            this.labelFileName.TabIndex = 21;
            this.labelFileName.Text = "file name";
            this.labelFileName.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
            // 
            // btnClearAltText
            // 
            this.btnClearAltText.BackColor = System.Drawing.Color.Yellow;
            this.btnClearAltText.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnClearAltText.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnClearAltText.Location = new System.Drawing.Point(680, 92);
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
            this.btnVerify.Location = new System.Drawing.Point(553, 92);
            this.btnVerify.Name = "btnVerify";
            this.btnVerify.Size = new System.Drawing.Size(105, 46);
            this.btnVerify.TabIndex = 23;
            this.btnVerify.Text = "Verify";
            this.btnVerify.UseVisualStyleBackColor = false;
            // 
            // btnAssign
            // 
            this.btnAssign.BackColor = System.Drawing.Color.Fuchsia;
            this.btnAssign.Cursor = System.Windows.Forms.Cursors.Hand;
            this.btnAssign.Font = new System.Drawing.Font("Segoe UI", 15.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.btnAssign.Location = new System.Drawing.Point(924, 92);
            this.btnAssign.Name = "btnAssign";
            this.btnAssign.Size = new System.Drawing.Size(105, 46);
            this.btnAssign.TabIndex = 24;
            this.btnAssign.Text = "Assign";
            this.btnAssign.UseVisualStyleBackColor = false;
            this.btnAssign.Click += new System.EventHandler(this.btnAssign_Click);
            // 
            // Form1
            // 
            this.ClientSize = new System.Drawing.Size(1052, 671);
            this.Controls.Add(this.btnAssign);
            this.Controls.Add(this.btnVerify);
            this.Controls.Add(this.btnClearAltText);
            this.Controls.Add(this.labelFileName);
            this.Controls.Add(this.buttonStop);
            this.Controls.Add(this.comboBoxWorksheetNames);
            this.Controls.Add(this.labelCollisions);
            this.Controls.Add(this.listBoxCollisions);
            this.Controls.Add(this.picBoxPanel);
            this.Controls.Add(this.labelNewImage);
            this.Controls.Add(this.labelOldImage);
            this.Controls.Add(this.lblExcel);
            this.Controls.Add(this.lblWord);
            this.Controls.Add(this.lblSheet);
            this.Controls.Add(this.txtExcelApp);
            this.Controls.Add(this.txtWordApp);
            this.Controls.Add(this.btnUpload);
            this.Controls.Add(this.btnBrowseExcel);
            this.Controls.Add(this.btnBrowseWord);
            this.Name = "Form1";
            this.Text = "Excel to GitHub Uploader";
            this.Load += new System.EventHandler(this.Form1_Load_1);
            ((System.ComponentModel.ISupportInitialize)(this.picBoxOldImage)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.picBoxNewImage)).EndInit();
            this.picBoxPanel.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox txtExcelApp;
        private System.Windows.Forms.TextBox txtWordApp;
        private System.Windows.Forms.Button btnUpload;
        private System.Windows.Forms.Label lblExcel;
        private System.Windows.Forms.Label lblWord;
        private System.Windows.Forms.Label lblSheet;
        private System.Windows.Forms.Button btnBrowseExcel;
        private System.Windows.Forms.Button btnBrowseWord;
        private System.Windows.Forms.PictureBox picBoxOldImage;
        private System.Windows.Forms.Label labelOldImage;
        private System.Windows.Forms.Label labelNewImage;
        private System.Windows.Forms.PictureBox picBoxNewImage;
        private System.Windows.Forms.Panel picBoxPanel;
        private System.Windows.Forms.ListBox listBoxCollisions;
        private System.Windows.Forms.Label labelCollisions;
        private System.Windows.Forms.ComboBox comboBoxWorksheetNames;
        private System.Windows.Forms.Button buttonStop;
        private System.Windows.Forms.Label labelFileName;
        private System.Windows.Forms.Button btnClearAltText;
        private System.Windows.Forms.Button btnVerify;
        private System.Windows.Forms.Button btnAssign;
    }
}
