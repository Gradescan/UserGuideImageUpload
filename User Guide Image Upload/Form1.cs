using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelWordImageUploader
{
    public partial class Form1 : Form
    {
        //-----------------------------------------------------------------------------------------
        private int COL_IMAGE_NAME;
        private int COL_ERRORS;
        private int COL_FILENAME;
        private int COL_TITLE;
        private int COL_COLOR;
        private int COL_TRANSFORM;
        private int COL_ICON;
        private int COL_ALT_TEXT_ID;
        private int COL_ALT_TEXT;
        private int COL_MAX_HEIGHT;

        //-----------------------------------------------------------------------------------------
        private string GitHubToken;
        private bool Stop = false;
        //-----------------------------------------------------------------------------------------
        // Maps sheet name -> HashSet of allowed (writable) column numbers
        private static readonly Dictionary<string, HashSet<int>> WritableColumnsBySheet = new Dictionary<string, HashSet<int>>();
        //-----------------------------------------------------------------------------------------
        public Form1()
        {
            System.Net.ServicePointManager.SecurityProtocol = System.Net.SecurityProtocolType.Tls12;

            InitializeComponent();
            StartPosition = FormStartPosition.WindowsDefaultLocation;

            GitHubToken = Environment.GetEnvironmentVariable("GITHUB_TOKEN");
            if (string.IsNullOrWhiteSpace(GitHubToken))
            {
                MessageBox.Show("Environment variable GITHUB_TOKEN is not set.");
                Environment.Exit(1);
            }

            comboBoxWorksheetNames.SelectedIndex = 1;
        }
        //-----------------------------------------------------------------------------------------
        private void Form1_Load(object sender, EventArgs e)
        {
            // Optional logic when the form loads
        }
        //-----------------------------------------------------------------------------------------
        private async void btnUpload_Click(object sender, EventArgs e)
        {
            Stop = false;

            string wordAppPath = txtWordApp.Text.Trim();
            string excelAppPath = txtExcelApp.Text.Trim();
            string sheetName = comboBoxWorksheetNames.Text.Trim();

            if (IsWorkbookOpen(excelAppPath))
            {
                MessageBox.Show("The Excel file is currently open. Please save and close the file.");
                return;
            }

            ZipArchive wordAppAsZip = null;
            ExcelPackage excelPackage = null;

            listBoxCollisions.Items.Clear();

            try
            {
                wordAppAsZip = ZipFile.OpenRead(wordAppPath);
                excelPackage = new ExcelPackage(new FileInfo(excelAppPath));
                var worksheet = excelPackage.Workbook.Worksheets[sheetName];

                COL_IMAGE_NAME = GetColumnNumberByHeaderTitle(excelPackage.Workbook, sheetName, "Image Name");
                COL_ERRORS = GetColumnNumberByHeaderTitle(excelPackage.Workbook, sheetName, "Errors");
                COL_FILENAME = GetColumnNumberByHeaderTitle(excelPackage.Workbook, sheetName, "File Name");
                COL_TITLE = GetColumnNumberByHeaderTitle(excelPackage.Workbook, sheetName, "Title");
                COL_ALT_TEXT_ID = GetColumnNumberByHeaderTitle(excelPackage.Workbook, sheetName, "AltTextId");
                COL_ICON = GetColumnNumberByHeaderTitle(excelPackage.Workbook, sheetName, "Icon");
                COL_COLOR = GetColumnNumberByHeaderTitle(excelPackage.Workbook, sheetName, "Color");
                COL_TRANSFORM = GetColumnNumberByHeaderTitle(excelPackage.Workbook, sheetName, "Transform");
                COL_ALT_TEXT = GetColumnNumberByHeaderTitle(excelPackage.Workbook, sheetName, "Alt Text");
                COL_MAX_HEIGHT = GetColumnNumberByHeaderTitle(excelPackage.Workbook, sheetName, "Height Max");

                // safety check - identify writeable columns
                WritableColumnsBySheet.Clear();
                WritableColumnsBySheet.Add(sheetName, new HashSet<int> { COL_ALT_TEXT, COL_ERRORS });

                for (int row = 2; row < 1000; row++)        // first row is header
                {
                    if (Stop)
                    {
                        excelPackage.Save();     // background cell colors
                        MessageBox.Show("Stopped");
                        return;
                    }

                    SafeSetCellValue(worksheet, row, COL_ERRORS, "");

                    // required. stop if not found
                    string sourceImageName = worksheet.Cells[row, COL_IMAGE_NAME].Text;
                    if (string.IsNullOrWhiteSpace(sourceImageName))
                        break;

                    // required. stop if not found
                    string altTextId = worksheet.Cells[row, COL_ALT_TEXT_ID].Text;
                    if (string.IsNullOrWhiteSpace(altTextId))
                        break;



                    string title = worksheet.Cells[row, COL_TITLE].Text;

                    // upload to repo only images - not icons
                    string icon = worksheet.Cells[row, COL_ICON].Text;

                    if (false == string.IsNullOrEmpty(icon))
                    {
                        // this is an icon - no image in repo required.
                        string colour = worksheet.Cells[row, COL_COLOR].Text;
                        string transform = worksheet.Cells[row, COL_TRANSFORM].Text;

                        // record the Alt Text to be written (by another process) to the User Guide
                        string altTextValue = GetAltText(icon, title, colour, transform, altTextId);
                        SafeSetCellValue(worksheet, row, COL_ALT_TEXT, altTextValue);
                        continue;
                    }

                    // upload image to repo

                    string destFileName = worksheet.Cells[row, COL_FILENAME].Text;
                    if (string.IsNullOrWhiteSpace(destFileName))
                        continue;

                    string destAltTextId = worksheet.Cells[row, COL_ALT_TEXT_ID].Text;
                    if (string.IsNullOrWhiteSpace(destAltTextId))
                        continue;

                    // ensures no duplicate file names
                    destFileName += "-" + worksheet.Cells[row, COL_ALT_TEXT_ID].Text;

                    BeginInvoke((Action)(() => labelNewImage.Text = sourceImageName));

                    string internalPath = "word/media/" + sourceImageName;
                    ZipArchiveEntry wordMediaImage = wordAppAsZip.GetEntry(internalPath);
                    if (wordMediaImage == null)
                    {
                        Console.WriteLine("Image not found: " + sourceImageName);
                        continue;
                    }

                    byte[] newImageBytes;
                    using (Stream sourceStream = wordMediaImage.Open())
                    {
                        using (MemoryStream ms = new MemoryStream())
                        {
                            sourceStream.CopyTo(ms);
                            newImageBytes = ms.ToArray();
                        }
                    }
                    int newHash = newImageBytes.Aggregate(17, (current, b) => current * 31 + b);

                    Bitmap bitmap;
                    using (var ms = new MemoryStream(newImageBytes))
                    {
                        bitmap = new Bitmap(ms);
                    }
                    picBoxNewImage.Image = bitmap;
                    // show the media image
                    BeginInvoke((Action)(() => labelFileName.Text = sourceImageName));

                    string sha = string.Empty;

                    string objinfo = await GetFileRepoAsync(destFileName).ConfigureAwait(false);
                    if (objinfo != null)
                    {
                        dynamic obj = JsonConvert.DeserializeObject(objinfo);

                        // lookup the existing image
                        string html_url = (string)obj.html_url;

                        // record the Alt Text to be written (by another process) to the User Guide
                        //string altTextValue = GetAltText(html_url, title, altTextId);
                        //SafeSetCellValue(worksheet, row, COL_ALT_TEXT, altTextValue);

                        string base64 = (string)obj.content;
                        base64 = base64.Replace("\n", "").Replace("\r", ""); // GitHub adds newlines to base64 output

                        byte[] oldImageBytes = Convert.FromBase64String(base64);
                        int oldHash = oldImageBytes.Aggregate(17, (current, b) => current * 31 + b);

                        using (Stream sourceStream = wordMediaImage.Open())
                        {
                            using (MemoryStream ms = new MemoryStream())
                            {
                                sourceStream.CopyTo(ms);
                                oldImageBytes = ms.ToArray();
                            }
                        }

                        using (var ms = new MemoryStream(oldImageBytes))
                        {
                            bitmap = new Bitmap(ms);
                        }
                        // show the image in the repo
                        picBoxOldImage.Image = bitmap;
                        picBoxPanel.BackColor = Color.PaleGreen;

                        sha = (string)obj.sha;
                        System.Diagnostics.Debug.WriteLine("SHA: " + sha);

                        if (newHash != oldHash)
                        {
                            //worksheet.Cells[row, 3].Value = "Error";
                            SafeSetCellValue(worksheet, row, COL_ERRORS, "Error");
                            BeginInvoke((Action)(() => listBoxCollisions.Items.Add(destFileName)));
                            picBoxPanel.BackColor = Color.Red;
                            MessageBox.Show("MisMatch: " + destFileName);
                        }
                    }
                    else
                    {
                        picBoxOldImage.Image = null;
                        picBoxPanel.BackColor = SystemColors.Control;
                    }
                    // push the image to the repo
                    bool result = PutFileRepoAsync(destFileName, sha, newImageBytes).GetAwaiter().GetResult();

                    if (result)
                    {
                        objinfo = await GetFileRepoAsync(destFileName).ConfigureAwait(false);
                        if (objinfo != null)
                        {
                            dynamic obj = JsonConvert.DeserializeObject(objinfo);

                            if (string.IsNullOrEmpty(icon))
                            {
                                string html_url = (string)obj.html_url;
                                string max_height = worksheet.Cells[row, COL_MAX_HEIGHT].Text;
                                // record the Alt Text to be written (by another process) to the User Guide
                                string altTextValue = GetAltText(html_url, title, max_height, altTextId);
                                SafeSetCellValue(worksheet, row, COL_ALT_TEXT, altTextValue);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("Upload failed: " + destFileName);
                        MessageBox.Show("Upload failed: " + destFileName);
                    }
                }
                excelPackage.Save();     // background cell colors

                Console.WriteLine("Upload process completed.");
                MessageBox.Show("Upload process completed.");
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                if (wordAppAsZip != null) wordAppAsZip.Dispose();
                if (excelPackage != null) excelPackage.Dispose();
            }
        }
        //-----------------------------------------------------------------------------------------
        private async Task<byte[]> DownloadFileFromGitHubAsync(string owner, string repo, string path)
        {
            string apiUrl = $"https://api.github.com/repos/{owner}/{repo}/contents/{Uri.EscapeDataString(path)}";

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("UploaderApp", "1.0"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("token", GitHubToken);

                try
                {
                    HttpResponseMessage response = await client.GetAsync(apiUrl).ConfigureAwait(false);
                    string json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    if (!response.IsSuccessStatusCode)
                    {
                        System.Diagnostics.Debug.WriteLine($"GitHub error: {response.StatusCode}");
                        System.Diagnostics.Debug.WriteLine("Response: " + json);
                        return null;
                    }

                    dynamic obj = JsonConvert.DeserializeObject(json);
                    string base64 = (string)obj.content;
                    base64 = base64.Replace("\n", "").Replace("\r", ""); // GitHub adds newlines to base64 output

                    byte[] contentBytes = Convert.FromBase64String(base64);
                    return contentBytes;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("DownloadFileFromGitHubAsync error: " + ex.Message);
                    return null;
                }
            }
        }
        //-----------------------------------------------------------------------------------------
        private async Task<bool> PutFileRepoAsync(string destFileName, string sha, byte[] imageBytes)
        {
            try
            {
                string fileName = $"{destFileName}.png";

                string url = $"https://api.github.com/repos/Gradescan/images/contents/{Uri.EscapeDataString(fileName)}";
                System.Diagnostics.Debug.WriteLine("PutFileRepoAsync URL: " + url);

                string base64 = Convert.ToBase64String(imageBytes);

                var payload = new
                {
                    message = "Auto upload " + fileName,
                    content = base64,
                    sha = sha, // <- only populated if the file exists
                    committer = new { name = "UploaderBot", email = "uploader@gradescan.org" },
                    author = new { name = "UploaderBot", email = "uploader@gradescan.org" }
                };

                string json = JsonConvert.SerializeObject(payload);

                System.Diagnostics.Debug.WriteLine("PutFileRepoAsync URL: " + url);

                using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(15) })
                {
                    client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("UploaderApp", "1.0"));
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("token", GitHubToken);

                    HttpContent httpContent = new StringContent(json, Encoding.UTF8, "application/json");
                    HttpResponseMessage response = await client.PutAsync(url, httpContent).ConfigureAwait(false);
                    string result = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    System.Diagnostics.Debug.WriteLine("IsSuccessStatusCode: " + response.IsSuccessStatusCode);
                    System.Diagnostics.Debug.WriteLine("ReasonPhrase: " + response.ReasonPhrase);

                    return response.IsSuccessStatusCode;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("PutFileRepoAsync error: " + ex.Message);
                return false;
            }
        }
        //-----------------------------------------------------------------------------------------
        private async Task<string> GetFileRepoAsync(string destFileName)
        {
            string fileName = $"{destFileName}.png";
            string url = $"https://api.github.com/repos/Gradescan/images/contents/{Uri.EscapeDataString(fileName)}";
            System.Diagnostics.Debug.WriteLine("GetFileRepoAsync URL: " + url);

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("UploaderApp", "1.0"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("token", GitHubToken);

                try
                {
                    HttpResponseMessage response = await client.GetAsync(url).ConfigureAwait(false);

                    string json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    if (!response.IsSuccessStatusCode)
                    {
                        System.Diagnostics.Debug.WriteLine("SHA lookup response: " + json);
                        return null;
                    }

                    return json;
                    //dynamic obj = JsonConvert.DeserializeObject(json);
                    //string sha = (string)obj.sha;
                    //System.Diagnostics.Debug.WriteLine("SHA: " + sha);
                    //return sha;

                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine("GetFileShaAsync error: " + ex.Message);
                    return null;
                }
            }
        }
        //-----------------------------------------------------------------------------------------
        public static bool IsWorkbookOpen(string filePath)
        {
            try
            {
                using (FileStream stream = new FileStream(filePath, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
                {
                    // File is not locked
                    return false;
                }
            }
            catch (IOException)
            {
                // File is locked (likely open in Excel)
                return true;
            }
        }
        //-----------------------------------------------------------------------------------------
        public static int GetColumnNumberByHeaderTitle(ExcelWorkbook workbook, string worksheetName, string columnTitle)
        {
            var worksheet = workbook.Worksheets[worksheetName];
            if (worksheet == null)
                throw new ArgumentException($"Worksheet '{worksheetName}' not found.");

            int startColumn = worksheet.Dimension.Start.Column;
            int endColumn = worksheet.Dimension.End.Column;

            for (int col = startColumn; col <= endColumn; col++)
            {
                var cellValue = worksheet.Cells[1, col].Text?.Trim();
                if (string.Equals(cellValue, columnTitle, StringComparison.OrdinalIgnoreCase))
                {
                    return col;
                }
            }

            throw new ArgumentException($"Column with title '{columnTitle}' not found in worksheet '{worksheetName}'.");
        }
        //-----------------------------------------------------------------------------------------
        public static void SafeSetCellValue(ExcelWorksheet worksheet, int row, int column, object value)
        {
            string sheetName = worksheet.Name;

            if (!WritableColumnsBySheet.TryGetValue(sheetName, out var allowedColumns))
                throw new InvalidOperationException($"Write rules for worksheet '{sheetName}' are not defined.");

            if (!allowedColumns.Contains(column))
                throw new UnauthorizedAccessException($"Write to column {column} is not permitted in worksheet '{sheetName}'.");

            worksheet.Cells[row, column].Value = value;
        }
        //-----------------------------------------------------------------------------------------
        private void btnBrowseExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "Select Excel File";
            dlg.Filter = "Excel Files (*.xlsx)|*.xlsx";
            dlg.InitialDirectory = Path.GetDirectoryName(txtExcelApp.Text);
            if (dlg.ShowDialog() == DialogResult.OK)
                txtExcelApp.Text = dlg.FileName;
        }
        //-----------------------------------------------------------------------------------------
        private void btnBrowseWord_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "Select Word File";
            dlg.Filter = "Word Files (*.docm)|*.docm";
            dlg.InitialDirectory = Path.GetDirectoryName(txtExcelApp.Text);
            if (dlg.ShowDialog() == DialogResult.OK)
                txtWordApp.Text = dlg.FileName;
        }
        //-----------------------------------------------------------------------------------------
        private void Form1_Load_1(object sender, EventArgs e)
        {

        }
        //-----------------------------------------------------------------------------------------
        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void txtExcel_TextChanged(object sender, EventArgs e)
        {

        }

        private void lblExcel_Click(object sender, EventArgs e)
        {

        }

        private void buttonStop_Click(object sender, EventArgs e)
        {
            Stop = true;
        }
        //-----------------------------------------------------------------------------------------
        //<img src="https://github.com/Gradescan/images/blob/main/question-edit-answers-ordered-222.png?raw=true" style="max-height: 200px; width: auto;" alt="Question Edit">

        private string GetAltText(string html_url, string title, string max_height, string altTextId)
        {
            string _maxheight = (string.IsNullOrEmpty(max_height) ? "" : $@"style=""max-height: {max_height}px; width: auto; "" ");
            string alttext =
        $@"<span style=""font-size: 18px;"" title=""{title} "">
    <img src=""{html_url}?raw=true"" {_maxheight}/>
</span> 
[{altTextId}]";

            // Replace two double quotes ("") with one double quote (")
            return alttext.Replace("\"\"", "\"");
        }
        //-----------------------------------------------------------------------------------------
        private string GetAltText(string icon, string title, string colour, string transform, string altTextId)
        {
            string _class = string.IsNullOrEmpty(transform) ? string.Empty : $@"class: ""{transform}"" ";
            string _style = (string.IsNullOrEmpty(colour) ? "" : $"color: {colour}; ")
                          + (string.IsNullOrEmpty(transform) ? "" : "display: inline-block;");
            string alttext =
        $@"<span style=""font-size: 20px; {_style.Trim()}"" {_class} title=""{title} "">
  {icon}
</span> 
[{altTextId}]";

            // Replace two double quotes ("") with one double quote (")
            return alttext.Replace("\"\"", "\"");
        }
        //-----------------------------------------------------------------------------------------
        private void btnAssign_Click(object sender, EventArgs e)
        {
            // Prompt user for starting number
            string input = Microsoft.VisualBasic.Interaction.InputBox("Enter starting number (e.g., 1):", "Start Number", "1");
            if (!int.TryParse(input, out int counter) || counter < 0)
            {
                MessageBox.Show("Invalid number entered.");
                return;
            }

            int largestNumber = counter;

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                string wordPath = txtWordApp.Text.Trim();
                wordApp = new Word.Application();
                doc = wordApp.Documents.Open(wordPath);

                Regex pattern = new Regex(@"\[(\d{3})\]");

                // Process InlineShapes
                foreach (Word.InlineShape shape in doc.InlineShapes)
                {
                    string altText = shape.AlternativeText ?? "";

                    Match match = pattern.Match(altText);
                    if (match.Success)
                    {
                        // Extract existing number and update largestNumber if greater
                        if (int.TryParse(match.Groups[1].Value, out int foundNumber))
                        {
                            if (foundNumber >= largestNumber)
                                largestNumber = foundNumber + 1;
                        }
                    }
                    else
                    {
                        // Assign new number
                        shape.AlternativeText = altText + "[" + counter.ToString("D3") + "]";
                        counter++;
                        largestNumber = counter;
                    }
                }

                foreach (Word.Shape shape in doc.Shapes)
                {
                    string altText = shape.AlternativeText;
                    if (!pattern.IsMatch(altText))
                    {
                        string newText = altText + "\n[" + counter.ToString("D3") + "]";
                        shape.AlternativeText = newText;
                        counter++;
                    }
                }
                doc.Save();
                MessageBox.Show("AltText update completed. Last value used or found = " + (largestNumber - 1).ToString());
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                // Ensure file is unlocked by closing document and quitting Word
                if (doc != null)
                {
                    doc.Close(false); // false = don't prompt to save again
                    Marshal.ReleaseComObject(doc);
                    doc = null;
                }

                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                    wordApp = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }
        //-----------------------------------------------------------------------------------------
        private void btnClearAltText_Click(object sender, EventArgs e)
        {
            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                string wordPath = txtWordApp.Text.Trim();
                wordApp = new Word.Application();
                doc = wordApp.Documents.Open(wordPath);

                int clearedCount = 0;
                int convertedCount = 0;

                // Clear InlineShapes Alt Text
                foreach (Word.InlineShape shape in doc.InlineShapes)
                {
                    if (!string.IsNullOrEmpty(shape.AlternativeText))
                    {
                        shape.AlternativeText = string.Empty;
                        clearedCount++;
                    }
                }

                // Clear floating Shapes Alt Text
                foreach (Word.Shape shape in doc.Shapes)
                {
                    // Check if it's not already inline (i.e., floating)
                    if (shape.WrapFormat.Type != Word.WdWrapType.wdWrapInline)
                    {
                        // Convert to InlineShape
                        shape.ConvertToInlineShape();
                        convertedCount++;
                    }
                    if (!string.IsNullOrEmpty(shape.AlternativeText))
                    {
                        shape.AlternativeText = string.Empty;
                        clearedCount++;
                    }
                }

                doc.Save();
                MessageBox.Show($"Cleared AltText from {clearedCount} shapes.  Converted {convertedCount} shapes.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                // Ensure file is unlocked by closing document and quitting Word
                if (doc != null)
                {
                    doc.Close(false); // false = don't prompt to save again
                    Marshal.ReleaseComObject(doc);
                    doc = null;
                }

                if (wordApp != null)
                {
                    wordApp.Quit();
                    Marshal.ReleaseComObject(wordApp);
                    wordApp = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        private void listBoxCollisions_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void comboBoxWorksheetNames_SelectedIndexChanged(object sender, EventArgs e)
        {

        }
        //-----------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------
    }
    //-----------------------------------------------------------------------------------------
}
