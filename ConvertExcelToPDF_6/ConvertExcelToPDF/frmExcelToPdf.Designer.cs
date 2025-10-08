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
            SuspendLayout();
            // 
            // txtExcelFolder
            // 
            txtExcelFolder.Location = new Point(120, 15);
            txtExcelFolder.Name = "txtExcelFolder";
            txtExcelFolder.Size = new Size(400, 23);
            txtExcelFolder.TabIndex = 0;
            // 
            // btnChooseExcelFolder
            // 
            btnChooseExcelFolder.Location = new Point(530, 15);
            btnChooseExcelFolder.Name = "btnChooseExcelFolder";
            btnChooseExcelFolder.Size = new Size(75, 23);
            btnChooseExcelFolder.TabIndex = 1;
            btnChooseExcelFolder.Text = "Browse...";
            btnChooseExcelFolder.UseVisualStyleBackColor = true;
            btnChooseExcelFolder.Click += btnChooseExcelFolder_Click;
            // 
            // txtPdfFolder
            // 
            txtPdfFolder.Location = new Point(120, 55);
            txtPdfFolder.Name = "txtPdfFolder";
            txtPdfFolder.Size = new Size(400, 23);
            txtPdfFolder.TabIndex = 2;
            // 
            // btnChoosePdfFolder
            // 
            btnChoosePdfFolder.Location = new Point(530, 55);
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
            btnConvert.Location = new Point(250, 110);
            btnConvert.Name = "btnConvert";
            btnConvert.Size = new Size(150, 40);
            btnConvert.TabIndex = 4;
            btnConvert.Text = "Convert to PDF";
            btnConvert.UseVisualStyleBackColor = true;
            btnConvert.Click += btnConvert_Click;
            // 
            // lblExcel
            // 
            lblExcel.AutoSize = true;
            lblExcel.Location = new Point(20, 23);
            lblExcel.Name = "lblExcel";
            lblExcel.Size = new Size(72, 15);
            lblExcel.TabIndex = 5;
            lblExcel.Text = "Excel Folder:";
            // 
            // lblPdf
            // 
            lblPdf.AutoSize = true;
            lblPdf.Location = new Point(20, 63);
            lblPdf.Name = "lblPdf";
            lblPdf.Size = new Size(72, 15);
            lblPdf.TabIndex = 6;
            lblPdf.Text = "PDF Output:";
            // 
            // button1
            // 
            button1.Location = new Point(559, 180);
            button1.Name = "button1";
            button1.Size = new Size(8, 8);
            button1.TabIndex = 7;
            button1.Text = "button1";
            button1.UseVisualStyleBackColor = true;
            // 
            // frmExcelToPdf
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(800, 450);
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
    }
}