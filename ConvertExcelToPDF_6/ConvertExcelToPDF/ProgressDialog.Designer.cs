namespace ConvertExcelToPDF
{
    partial class ProgressDialog
    {
        private System.ComponentModel.IContainer components = null;
        public ProgressBar progressBarConversion;
        public Button btnCancel;
        public Label lblCurrentFile;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.progressBarConversion = new System.Windows.Forms.ProgressBar();
            this.btnCancel = new System.Windows.Forms.Button();
            this.lblCurrentFile = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // progressBarConversion
            // 
            this.progressBarConversion.Location = new System.Drawing.Point(12, 40);
            this.progressBarConversion.Name = "progressBarConversion";
            this.progressBarConversion.Size = new System.Drawing.Size(260, 23);
            this.progressBarConversion.Style = System.Windows.Forms.ProgressBarStyle.Continuous;
            // 
            // btnCancel
            // 
            this.btnCancel.Location = new System.Drawing.Point(197, 69);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.Text = "Cancel";
            this.btnCancel.UseVisualStyleBackColor = true;
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lblCurrentFile
            // 
            this.lblCurrentFile.Location = new System.Drawing.Point(12, 10);
            this.lblCurrentFile.Name = "lblCurrentFile";
            this.lblCurrentFile.Size = new System.Drawing.Size(260, 20);
            this.lblCurrentFile.Text = "Starting conversion...";
            this.lblCurrentFile.AutoSize = false;
            // 
            // ProgressDialog
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(284, 104);
            this.Controls.Add(this.lblCurrentFile);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.progressBarConversion);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "ProgressDialog";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Converting Files...";
            this.ResumeLayout(false);
        }
    }
}