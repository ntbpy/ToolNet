using Spire.Pdf;
using Spire.Pdf.Graphics;
using Spire.Xls;
using Spire.Xls.AI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using XlsFileFormat = Spire.Xls.FileFormat;
using Excel = Microsoft.Office.Interop.Excel;

namespace ConvertExcelToPDF
{
    public partial class frmExcelToPdf : Form
    {
        private readonly string _tempFolder = Path.Combine(Path.GetTempPath(), "ExcelToPdfTemp");
        private string _logFilePath;

        public frmExcelToPdf()
        {
            InitializeComponent();
        }

        private void btnChooseExcelFolder_Click(object sender, EventArgs e)
        {
            SelectFolder(txtExcelFolder);
        }

        private void btnChoosePdfFolder_Click(object sender, EventArgs e)
        {
            SelectFolder(txtPdfFolder);
        }

        private void SelectFolder(TextBox textBox)
        {
            using var folderDialog = new FolderBrowserDialog();
            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                textBox.Text = folderDialog.SelectedPath;
            }
        }

        private async void btnConvert_Click(object sender, EventArgs e)
        {
            try
            {
                string excelFolder = txtDesExcelFolder.Text.Trim();
                string pdfFolder = txtPdfFolder.Text.Trim();

                if (!ValidateFolders(excelFolder, pdfFolder))
                {
                    return;
                }

                // Lazy enumeration of Excel files
                var excelFiles = Directory.EnumerateFiles(excelFolder)
                    .Where(file => file.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                                  file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    .GroupBy(file => Path.GetFileNameWithoutExtension(file), StringComparer.OrdinalIgnoreCase)
                    .Select(group => group.OrderByDescending(file => file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)).First())
                    .ToArray();

                if (excelFiles.Length == 0)
                {
                    MessageBox.Show("No Excel files found in the selected folder.", "Information",
                        MessageBoxButtons.OK, MessageBoxIcon.Information);
                    return;
                }

                // Initialize log file
                _logFilePath = Path.Combine(pdfFolder, $"ConversionLog_{DateTime.Now:yyyyMMdd_HHmmss}.txt");
                await LogAsync($"Conversion started at {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                await LogAsync($"Input folder: {excelFolder}");
                await LogAsync($"Output folder: {pdfFolder}");
                await LogAsync($"Found {excelFiles.Length} Excel files to process.");
                await LogAsync($"Files: {string.Join(", ", excelFiles.Select(Path.GetFileName))}");

                // Initialize cancellation token and show progress dialog
                var cts = new CancellationTokenSource();
                var progressDialog = new ProgressDialog(excelFiles.Length, cts);
                progressDialog.Show(); // Non-modal dialog
                ToggleUI(false);

                Directory.CreateDirectory(_tempFolder);
                var (success, failed, failedFiles, wasCancelled) = await ConvertFilesAsync(excelFiles, pdfFolder, cts.Token, progressDialog);
                ToggleUI(true);

                // Close progress dialog
                progressDialog.CloseDialog();

                // Log summary
                await LogAsync($"Conversion finished at {DateTime.Now:yyyy-MM-dd HH:mm:ss}");
                await LogAsync($"Summary: Succeeded: {success}, Failed: {failed}");
                if (wasCancelled)
                {
                    await LogAsync("Operation was cancelled by user.");
                }
                if (failedFiles.Any())
                {
                    await LogAsync("Failed files:\n" + string.Join("\n", failedFiles.Select(Path.GetFileName)));
                }
                await LogAsync("----------------------------------------");

                // Build result message
                string message = $"Conversion finished.\n\nSucceeded: {success}\nFailed: {failed}";
                if (failed > 0 && failedFiles.Any())
                {
                    message += "\n\nFiles that failed to convert:\n" + string.Join("\n", failedFiles.Select(Path.GetFileName));
                }
                if (wasCancelled)
                {
                    message += "\n\nOperation was cancelled by user.";
                }
                message += $"\n\nDetails saved to log file: {Path.GetFileName(_logFilePath)}";
                MessageBox.Show(message, wasCancelled ? "Cancelled" : "Done", MessageBoxButtons.OK,
                    wasCancelled ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            }
            catch (Exception ex)
            {
                await LogAsync($"Error in conversion process: {ex.Message}\nStackTrace: {ex.StackTrace}");
                MessageBox.Show($"An error occurred: {ex.Message}\n\nDetails saved to log file: {Path.GetFileName(_logFilePath)}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                CleanupTempFiles();
                ToggleUI(true);
            }
        }

        private async Task LogAsync(string message)
        {
            try
            {
                using (var writer = new StreamWriter(_logFilePath, true))
                {
                    await writer.WriteLineAsync(message);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Logging error: {ex.Message}");
            }
        }

        private bool ValidateFolders(string excelFolder, string pdfFolder)
        {
            if (string.IsNullOrEmpty(excelFolder) || !Directory.Exists(excelFolder))
            {
                MessageBox.Show("Excel folder does not exist. Please select a valid folder.",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            if (string.IsNullOrEmpty(pdfFolder) || !Directory.Exists(pdfFolder))
            {
                MessageBox.Show("PDF output folder does not exist. Please select a valid folder.",
                    "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            // Check write permissions for output folder
            try
            {
                string testFile = Path.Combine(pdfFolder, "test_write_permissions.txt");
                File.WriteAllText(testFile, "test");
                File.Delete(testFile);
            }
            catch (Exception ex)
            {
                MessageBox.Show($"Cannot write to PDF output folder: {ex.Message}", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return false;
            }

            return true;
        }

        private void ToggleUI(bool enabled)
        {
            btnConvert.Enabled = enabled;
            btnChooseExcelFolder.Enabled = enabled;
            btnChoosePdfFolder.Enabled = enabled;
        }

        private async Task<(int success, int failed, List<string> failedFiles, bool wasCancelled)> ConvertFilesAsync(string[] excelFiles, string pdfFolder, CancellationToken cancellationToken, ProgressDialog progressDialog)
        {
            int success = 0;
            int failed = 0;
            var failedFiles = new List<string>();

            try
            {
                await Parallel.ForEachAsync(excelFiles, new ParallelOptions { MaxDegreeOfParallelism = Environment.ProcessorCount, CancellationToken = cancellationToken }, async (file, ct) =>
                {
                    try
                    {
                        // Update current file before processing
                        if (!progressDialog.IsDisposed)
                        {
                            progressDialog.Invoke(new Action(() => progressDialog.UpdateProgress(Path.GetFileName(file))));
                        }
                        await LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Started processing: {Path.GetFileName(file)}");

                        // Verify file exists and is accessible
                        if (!File.Exists(file))
                        {
                            throw new FileNotFoundException($"Excel file not found: {file}");
                        }

                        string pdfFileName = GetUniquePdfFileName(pdfFolder, Path.GetFileNameWithoutExtension(file));
                        await Task.Run(() => ExportExcelToPdf(file, pdfFileName, ct), ct);
                        Interlocked.Increment(ref success);
                        await LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Successfully converted: {Path.GetFileName(file)} to {Path.GetFileName(pdfFileName)}");

                        // Verify PDF was created
                        if (!File.Exists(pdfFileName))
                        {
                            throw new IOException($"PDF file was not created: {pdfFileName}");
                        }
                    }
                    catch (OperationCanceledException)
                    {
                        await LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Cancelled processing: {Path.GetFileName(file)}");
                        throw;
                    }
                    catch (Exception ex)
                    {
                        Interlocked.Increment(ref failed);
                        lock (failedFiles)
                        {
                            failedFiles.Add(file);
                        }
                        await LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Failed to convert: {Path.GetFileName(file)}\nError: {ex.Message}\nStackTrace: {ex.StackTrace}");
                    }
                });
            }
            catch (OperationCanceledException)
            {
                await LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Conversion cancelled by user.");
                return (success, failed, failedFiles, true);
            }

            return (success, failed, failedFiles, false);
        }

        private string GetUniquePdfFileName(string pdfFolder, string baseFileName)
        {
            string pdfFileName = Path.Combine(pdfFolder, $"{baseFileName}.pdf");
            if (!File.Exists(pdfFileName))
            {
                return pdfFileName;
            }

            int version = 1;
            string newFileName;
            do
            {
                newFileName = Path.Combine(pdfFolder, $"{baseFileName}_{version}.pdf");
                version++;
            } while (File.Exists(newFileName));

            return newFileName;
        }
        //private void ExportExcelToPdf(string excelPath, string pdfPath, CancellationToken cancellationToken)
        //{
        //    try
        //    {
        //        LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Loading Excel file: {Path.GetFileName(excelPath)}").GetAwaiter().GetResult();
        //        using var workbook = new Workbook();
        //        workbook.LoadFromFile(excelPath);
        //        LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Loaded Excel file: {Path.GetFileName(excelPath)}").GetAwaiter().GetResult();

        //        using var finalDoc = new PdfDocument();
        //        int sheetCount = workbook.Worksheets.Count;
        //        LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Processing {sheetCount} worksheets in {Path.GetFileName(excelPath)}").GetAwaiter().GetResult();

        //        foreach (Worksheet sheet in workbook.Worksheets)
        //        {
        //            cancellationToken.ThrowIfCancellationRequested();
        //            LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Processing worksheet: {sheet.Name}").GetAwaiter().GetResult();

        //            using var singleSheetWorkbook = new Workbook();
        //            singleSheetWorkbook.Worksheets.Clear();
        //            singleSheetWorkbook.Worksheets.AddCopy(sheet);

        //            var adjustedSheet = singleSheetWorkbook.Worksheets[0];
        //            adjustedSheet.PageSetup.FitToPagesWide = 1;
        //            adjustedSheet.PageSetup.FitToPagesTall = 0; 

        //            using var pdfStream = new MemoryStream();
        //            singleSheetWorkbook.SaveToStream(pdfStream, XlsFileFormat.PDF);
        //            pdfStream.Position = 0;
        //            LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Converted worksheet {sheet.Name} to PDF stream").GetAwaiter().GetResult();

        //            using var partDoc = new PdfDocument();
        //            partDoc.LoadFromStream(pdfStream);
        //            LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Loaded PDF stream for worksheet {sheet.Name}").GetAwaiter().GetResult();

        //            // Add sheet name as a header to each page in partDoc
        //            foreach (PdfPageBase page in partDoc.Pages)
        //            {
        //                PdfFont font = new PdfFont(PdfFontFamily.Helvetica, 12f);
        //                PdfBrush brush = PdfBrushes.Black;
        //                string headerText = $"Sheet: {sheet.Name}";
        //                // Draw sheet name at top-left (10, 10) with margin
        //                page.Canvas.DrawString(headerText, font, brush, new PointF(10, 10));
        //                LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Added sheet name '{sheet.Name}' to page {partDoc.Pages.IndexOf(page) + 1}").GetAwaiter().GetResult();
        //            }

        //            finalDoc.AppendPage(partDoc);
        //            LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Appended worksheet {sheet.Name} to final PDF").GetAwaiter().GetResult();
        //        }

        //        cancellationToken.ThrowIfCancellationRequested();
        //        LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Saving PDF to: {Path.GetFileName(pdfPath)}").GetAwaiter().GetResult();
        //        finalDoc.SaveToFile(pdfPath);
        //        LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Saved PDF to: {Path.GetFileName(pdfPath)}").GetAwaiter().GetResult();
        //    }
        //    catch (Exception ex)
        //    {
        //        throw new Exception($"Failed to convert {Path.GetFileName(excelPath)} to PDF", ex);
        //    }
        //}

        private void ExportExcelToPdf(string excelPath, string pdfPath, CancellationToken cancellationToken)
        {
            try
            {
                LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Starting conversion: {Path.GetFileName(excelPath)}").GetAwaiter().GetResult();


                // Load Excel
                using var workbook = new Workbook();
                workbook.LoadFromFile(excelPath);
                LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Loaded Excel workbook with {workbook.Worksheets.Count} sheets.").GetAwaiter().GetResult();

                if (workbook.Worksheets.Count == 1)
                {
                    workbook.SaveToFile(pdfPath, XlsFileFormat.PDF);
                    LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Saved single-sheet PDF: {Path.GetFileName(pdfPath)}").GetAwaiter().GetResult();
                    return;
                }

                using var finalDoc = new PdfDocument();
                var tempParts = new List<(PdfDocument doc, MemoryStream stream)>();

                int sheetIndex = 0;
                foreach (Worksheet sheet in workbook.Worksheets)
                {
                    cancellationToken.ThrowIfCancellationRequested();
                    sheetIndex++;
                    string sheetName = string.IsNullOrWhiteSpace(sheet.Name) ? $"Sheet{sheetIndex}" : sheet.Name;

                    LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Processing sheet: {sheetName}").GetAwaiter().GetResult();

                    using var singleSheetWorkbook = new Workbook();
                    singleSheetWorkbook.Worksheets.Clear();
                    singleSheetWorkbook.Worksheets.AddCopy(sheet);

                    var ws = singleSheetWorkbook.Worksheets[0];
                    ws.PageSetup.PrintArea = null;
                    ws.PageSetup.FitToPagesWide = 1;
                    ws.PageSetup.FitToPagesTall = 1;
                    ws.PageSetup.IsPrintHeadings = false;
                    ws.PageSetup.IsPrintGridlines = false;

                    var pdfStream = new MemoryStream();
                    singleSheetWorkbook.SaveToStream(pdfStream, XlsFileFormat.PDF);
                    pdfStream.Position = 0;

                    var partDoc = new PdfDocument(pdfStream);
                    tempParts.Add((partDoc, pdfStream));

                    finalDoc.AppendPage(partDoc);
                    LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Appended {sheetName} to final PDF.").GetAwaiter().GetResult();
                }


                cancellationToken.ThrowIfCancellationRequested();
                LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Saving merged PDF: {Path.GetFileName(pdfPath)}").GetAwaiter().GetResult();
                finalDoc.SaveToFile(pdfPath);
                LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Successfully saved: {pdfPath}").GetAwaiter().GetResult();

                foreach (var (doc, stream) in tempParts)
                {
                    doc.Dispose();
                    stream.Dispose();
                }
            }
            catch (OperationCanceledException)
            {
                LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Conversion cancelled for {Path.GetFileName(excelPath)}").GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] ERROR converting {Path.GetFileName(excelPath)}: {ex}").GetAwaiter().GetResult();
                throw new Exception($"Failed to convert {Path.GetFileName(excelPath)} to PDF. See log for details.", ex);
            }
        }


        private string ConvertXlsToXlsx_Interop(string sourcePath, string pdfPath)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;

            try
            {
                string baseDir = Path.GetDirectoryName(pdfPath);
                if (string.IsNullOrEmpty(baseDir))
                    baseDir = Path.GetTempPath();

                string outputDir = Path.Combine(baseDir, "Converted");
                if (!Directory.Exists(outputDir))
                    Directory.CreateDirectory(outputDir);

                string pdfFileName = Path.GetFileNameWithoutExtension(pdfPath);
                string destPath = Path.Combine(outputDir, $"{pdfFileName}_converted.xlsx");

                LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Converting {Path.GetFileName(sourcePath)} → {Path.GetFileName(destPath)} using Excel Interop...").GetAwaiter().GetResult();

                excelApp = new Excel.Application
                {
                    DisplayAlerts = false,
                    Visible = false,
                    ScreenUpdating = false
                };

                workbook = excelApp.Workbooks.Open(
                    sourcePath,
                    ReadOnly: false,
                    Editable: true,
                    IgnoreReadOnlyRecommended: true
                );

                workbook.SaveAs(
                    destPath,
                    Excel.XlFileFormat.xlOpenXMLWorkbook,
                    AccessMode: Excel.XlSaveAsAccessMode.xlNoChange
                );

                LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Conversion done: {destPath}").GetAwaiter().GetResult();

                return destPath;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to convert {Path.GetFileName(sourcePath)} to XLSX using Interop", ex);
            }
            finally
            {
                try
                {
                    if (workbook != null)
                    {
                        try
                        {
                            workbook.Close(false);
                        }
                        catch
                        {
                            // Có thể workbook đã tự đóng sau SaveAs
                        }
                        finally
                        {
                            System.Runtime.InteropServices.Marshal.FinalReleaseComObject(workbook);
                        }
                    }

                    if (excelApp != null)
                    {
                        excelApp.Quit();
                        System.Runtime.InteropServices.Marshal.FinalReleaseComObject(excelApp);
                    }
                }
                catch (Exception ex2)
                {
                    LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Cleanup warning: {ex2.Message}").GetAwaiter().GetResult();
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void ExportExcelToPdf_Interop(string excelPath, string pdfPath, CancellationToken cancellationToken)
        {
            Excel.Application excelApp = null;
            Excel.Workbook excelWorkbook = null;
            try
            {
                LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Using Excel Interop fallback for {Path.GetFileName(excelPath)}").GetAwaiter().GetResult();

                excelApp = new Excel.Application
                {
                    Visible = false,
                    ScreenUpdating = false,
                    DisplayAlerts = false
                };

                excelWorkbook = excelApp.Workbooks.Open(excelPath);
                cancellationToken.ThrowIfCancellationRequested();

                // Export to PDF directly via Excel
                excelWorkbook.ExportAsFixedFormat(Excel.XlFixedFormatType.xlTypePDF, pdfPath);

                LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Excel Interop exported PDF successfully: {Path.GetFileName(pdfPath)}").GetAwaiter().GetResult();
            }
            catch (Exception ex)
            {
                LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Interop Excel failed: {ex.Message}").GetAwaiter().GetResult();
                throw;
            }
            finally
            {
                if (excelWorkbook != null)
                {
                    excelWorkbook.Close(false);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelWorkbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void CleanupTempFiles()
        {
            try
            {
                if (Directory.Exists(_tempFolder))
                {
                    Directory.Delete(_tempFolder, true);
                    LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Cleaned up temporary folder: {_tempFolder}").GetAwaiter().GetResult();
                }
            }
            catch (Exception ex)
            {
                LogAsync($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Failed to clean up temporary folder: {ex.Message}").GetAwaiter().GetResult();
            }
        }

        private void btnChooseDescExcelFolder_Click(object sender, EventArgs e)
        {
            SelectFolder(txtDesExcelFolder);
        }

        private void bnConvertToXLSX_Click(object sender, EventArgs e)
        {
            string sourceExcelFolder = txtExcelFolder.Text.Trim();
            string destinationExcelFolder = txtDesExcelFolder.Text.Trim();

            if (!Directory.Exists(sourceExcelFolder))
            {
                MessageBox.Show("Source folder does not exist.", "Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
                return;
            }

            if (!Directory.Exists(destinationExcelFolder))
            {
                Directory.CreateDirectory(destinationExcelFolder);
            }

            var excelFiles = Directory.EnumerateFiles(sourceExcelFolder)
                .Where(file => file.EndsWith(".xls", StringComparison.OrdinalIgnoreCase) ||
                               file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                .GroupBy(file => Path.GetFileNameWithoutExtension(file), StringComparer.OrdinalIgnoreCase)
                .Select(group => group.OrderByDescending(file => file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase)).First())
                .ToArray();

            if (excelFiles.Length == 0)
            {
                MessageBox.Show("No Excel files found in the selected folder.", "Information",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Excel.Application excelApp = null;
            try
            {
                excelApp = new Excel.Application();
                excelApp.DisplayAlerts = false;

                foreach (var file in excelFiles)
                {
                    string fileName = Path.GetFileNameWithoutExtension(file);
                    string destFilePath = Path.Combine(destinationExcelFolder, fileName + ".xlsx");

                    if (file.EndsWith(".xls", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            Excel.Workbook workbook = excelApp.Workbooks.Open(file);
                            workbook.SaveAs(destFilePath, Excel.XlFileFormat.xlOpenXMLWorkbook);
                            workbook.Close();
                            LogAsync($"Converted .xls to .xlsx: {fileName}");
                        }
                        catch (Exception ex)
                        {
                            LogAsync($"Failed to convert {fileName}.xls: {ex.Message}");
                        }
                    }
                    else if (file.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            File.Copy(file, destFilePath, true);
                            LogAsync($"Copied .xlsx file: {fileName}");
                        }
                        catch (Exception ex)
                        {
                            LogAsync($"Failed to copy {fileName}.xlsx: {ex.Message}");
                        }
                    }
                }

                MessageBox.Show("Conversion process completed.", "Done",
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            finally
            {
                if (excelApp != null)
                {
                    excelApp.Quit();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(excelApp);
                }
            }
        }
    }
}