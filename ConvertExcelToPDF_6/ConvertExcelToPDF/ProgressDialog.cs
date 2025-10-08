using System;
using System.Windows.Forms;
using System.Threading;

namespace ConvertExcelToPDF
{
    public partial class ProgressDialog : Form
    {
        private readonly CancellationTokenSource _cts;

        public ProgressDialog(int maxFiles, CancellationTokenSource cts)
        {
            InitializeComponent();
            _cts = cts;
            progressBarConversion.Maximum = maxFiles;
            progressBarConversion.Minimum = 0;
            progressBarConversion.Value = 0;
            lblCurrentFile.Text = "Starting conversion...";
        }

        public void UpdateProgress(string fileName)
        {
            if (!IsDisposed && progressBarConversion.Value < progressBarConversion.Maximum)
            {
                progressBarConversion.Value++;
                lblCurrentFile.Text = $"Processing: {fileName}";
            }
        }

        public void CloseDialog()
        {
            if (!IsDisposed)
            {
                Invoke(new Action(() => Close()));
            }
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            _cts.Cancel();
            btnCancel.Enabled = false;
            lblCurrentFile.Text = "Cancelling...";
        }

        protected override void OnFormClosing(FormClosingEventArgs e)
        {
            base.OnFormClosing(e);
            _cts?.Dispose();
        }
    }
}