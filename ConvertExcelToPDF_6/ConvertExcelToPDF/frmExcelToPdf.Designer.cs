namespace ConvertExcelToPDF
{
    partial class frmExcelToPdf
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
            txtExcelFolder = new TextBox();
            btnChooseExcelFolder = new Button();
            txtPdfFolder = new TextBox();
            btnChoosePdfFolder = new Button();
            btnConvert = new Button();
            lblExcel = new Label();
            lblPdf = new Label();
            button1 = new Button();
            txtDesExcelFolder = new TextBox();
            btnChooseDescExcelFolder = new Button();
            label1 = new Label();
            bnConvertToXLSX = new Button();
            lstFailedFiles = new ListBox();
            bnRetryFailed = new Button();
            lblStatus = new Label();
            SuspendLayout();
            // 
            // txtExcelFolder
            // 
            txtExcelFolder.Location = new Point(153, 11);
            txtExcelFolder.Name = "txtExcelFolder";
            txtExcelFolder.Size = new Size(400, 23);
            txtExcelFolder.TabIndex = 0;
            // 
            // btnChooseExcelFolder
            // 
            btnChooseExcelFolder.Location = new Point(563, 11);
            btnChooseExcelFolder.Name = "btnChooseExcelFolder";
            btnChooseExcelFolder.Size = new Size(75, 23);
            btnChooseExcelFolder.TabIndex = 1;
            btnChooseExcelFolder.Text = "Browse...";
            btnChooseExcelFolder.UseVisualStyleBackColor = true;
            btnChooseExcelFolder.Click += btnChooseExcelFolder_Click;
            // 
            // txtPdfFolder
            // 
            txtPdfFolder.Location = new Point(153, 195);
            txtPdfFolder.Name = "txtPdfFolder";
            txtPdfFolder.Size = new Size(400, 23);
            txtPdfFolder.TabIndex = 2;
            // 
            // btnChoosePdfFolder
            // 
            btnChoosePdfFolder.Location = new Point(563, 195);
            btnChoosePdfFolder.Name = "btnChoosePdfFolder";
            btnChoosePdfFolder.Size = new Size(75, 23);
            btnChoosePdfFolder.TabIndex = 3;
            btnChoosePdfFolder.Text = "Browse...";
            btnChoosePdfFolder.UseVisualStyleBackColor = true;
            btnChoosePdfFolder.Click += btnChoosePdfFolder_Click;
            // 
            // btnConvert
            // 
            btnConvert.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            btnConvert.Location = new Point(655, 195);
            btnConvert.Name = "btnConvert";
            btnConvert.Size = new Size(150, 34);
            btnConvert.TabIndex = 4;
            btnConvert.Text = "Convert to PDF";
            btnConvert.UseVisualStyleBackColor = true;
            btnConvert.Click += btnConvert_Click;
            // 
            // lblExcel
            // 
            lblExcel.AutoSize = true;
            lblExcel.Location = new Point(3, 23);
            lblExcel.Name = "lblExcel";
            lblExcel.Size = new Size(111, 15);
            lblExcel.TabIndex = 5;
            lblExcel.Text = "Source Excel Folder:";
            // 
            // lblPdf
            // 
            lblPdf.AutoSize = true;
            lblPdf.Location = new Point(20, 207);
            lblPdf.Name = "lblPdf";
            lblPdf.Size = new Size(72, 15);
            lblPdf.TabIndex = 6;
            lblPdf.Text = "PDF Output:";
            // 
            // button1
            // 
            button1.Location = new Point(592, 156);
            button1.Name = "button1";
            button1.Size = new Size(8, 8);
            button1.TabIndex = 7;
            button1.Text = "button1";
            button1.UseVisualStyleBackColor = true;
            // 
            // txtDesExcelFolder
            // 
            txtDesExcelFolder.Location = new Point(153, 54);
            txtDesExcelFolder.Name = "txtDesExcelFolder";
            txtDesExcelFolder.Size = new Size(400, 23);
            txtDesExcelFolder.TabIndex = 8;
            // 
            // btnChooseDescExcelFolder
            // 
            btnChooseDescExcelFolder.Location = new Point(563, 54);
            btnChooseDescExcelFolder.Name = "btnChooseDescExcelFolder";
            btnChooseDescExcelFolder.Size = new Size(75, 23);
            btnChooseDescExcelFolder.TabIndex = 9;
            btnChooseDescExcelFolder.Text = "Browse...";
            btnChooseDescExcelFolder.UseVisualStyleBackColor = true;
            btnChooseDescExcelFolder.Click += btnChooseDescExcelFolder_Click;
            // 
            // label1
            // 
            label1.AutoSize = true;
            label1.Location = new Point(3, 66);
            label1.Name = "label1";
            label1.Size = new Size(124, 15);
            label1.TabIndex = 10;
            label1.Text = "Des. New Excel Folder:";
            // 
            // bnConvertToXLSX
            // 
            bnConvertToXLSX.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            bnConvertToXLSX.Location = new Point(655, 48);
            bnConvertToXLSX.Name = "bnConvertToXLSX";
            bnConvertToXLSX.Size = new Size(150, 33);
            bnConvertToXLSX.TabIndex = 11;
            bnConvertToXLSX.Text = "Convert to XLSX";
            bnConvertToXLSX.UseVisualStyleBackColor = true;
            bnConvertToXLSX.Click += bnConvertToXLSX_Click;
            // 
            // lstFailedFiles
            // 
            lstFailedFiles.FormattingEnabled = true;
            lstFailedFiles.ItemHeight = 15;
            lstFailedFiles.Location = new Point(153, 112);
            lstFailedFiles.Name = "lstFailedFiles";
            lstFailedFiles.Size = new Size(485, 64);
            lstFailedFiles.TabIndex = 12;
            // 
            // bnRetryFailed
            // 
            bnRetryFailed.Enabled = false;
            bnRetryFailed.Font = new Font("Segoe UI", 10F, FontStyle.Bold, GraphicsUnit.Point);
            bnRetryFailed.Location = new Point(655, 142);
            bnRetryFailed.Name = "bnRetryFailed";
            bnRetryFailed.Size = new Size(150, 34);
            bnRetryFailed.TabIndex = 13;
            bnRetryFailed.Text = "Retry Failed Files";
            bnRetryFailed.UseVisualStyleBackColor = true;
            bnRetryFailed.Click += bnRetryFailed_ClickAsync;
            // 
            // lblStatus
            // 
            lblStatus.AutoSize = true;
            lblStatus.Location = new Point(120, 91);
            lblStatus.Name = "lblStatus";
            lblStatus.Size = new Size(38, 15);
            lblStatus.TabIndex = 14;
            lblStatus.Text = "label2";
            // 
            // frmExcelToPdf
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(867, 296);
            Controls.Add(lblStatus);
            Controls.Add(bnRetryFailed);
            Controls.Add(lstFailedFiles);
            Controls.Add(bnConvertToXLSX);
            Controls.Add(txtDesExcelFolder);
            Controls.Add(btnChooseDescExcelFolder);
            Controls.Add(label1);
            Controls.Add(button1);
            Controls.Add(txtExcelFolder);
            Controls.Add(btnChooseExcelFolder);
            Controls.Add(txtPdfFolder);
            Controls.Add(btnChoosePdfFolder);
            Controls.Add(btnConvert);
            Controls.Add(lblExcel);
            Controls.Add(lblPdf);
            Name = "frmExcelToPdf";
            Text = "Excel to PDF Converter";
            ResumeLayout(false);
            PerformLayout();
        }

        #endregion

        private System.Windows.Forms.TextBox txtExcelFolder;
        private System.Windows.Forms.Button btnChooseExcelFolder;
        private System.Windows.Forms.TextBox txtPdfFolder;
        private System.Windows.Forms.Button btnChoosePdfFolder;
        private System.Windows.Forms.Button btnConvert;
        private System.Windows.Forms.Label lblExcel;
        private System.Windows.Forms.Label lblPdf;
        private Button button1;
        private TextBox txtDesExcelFolder;
        private Button btnChooseDescExcelFolder;
        private Label label1;
        private Button bnConvertToXLSX;
        private ListBox lstFailedFiles;
        private Button bnRetryFailed;
        private Label lblStatus;
    }
}