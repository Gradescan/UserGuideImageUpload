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
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;
using Word = Microsoft.Office.Interop.Word;

namespace ExcelWordImageUploader
{
    //-----------------------------------------------------------------------------------------
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
        private int COL_ALT_TEXT;
        private int COL_MAX_HEIGHT;

        //-----------------------------------------------------------------------------------------
        public class GitReadStatus
        {
            public string message { get; set; }
            public string documentation_url { get; set; }
            public string status { get; set; }
        }
        //-----------------------------------------------------------------------------------------
        public class GitHubFileInfo
        {
            public string sha { get; set; }
            public string url { get; set; }
            public string html_url { get; set; }
            public string content { get; set; }
            public string encoding { get; set; }
        }
        //-----------------------------------------------------------------------------------------
        private WorksheetItem[] worksheetItems = new[]
        {
            new WorksheetItem("Professional Edition", 1000),
            new WorksheetItem("Basic Edition", 2000),
            new WorksheetItem("Professional Forms", 3000),
            new WorksheetItem("Basic Forms", 4000),
        };
        //-----------------------------------------------------------------------------------------
        private const string SelectaUserGuide = "Select a User Guide";
        private List<(string name, string sha)> _cachedImageFilesInRepo = null;

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
            this.Load += Form1_Load;  // Ensures Form1_Load gets called
            StartPosition = FormStartPosition.WindowsDefaultLocation;

            GitHubToken = Environment.GetEnvironmentVariable("GITHUB_TOKEN");
            if (string.IsNullOrWhiteSpace(GitHubToken))
            {
                MessageBox.Show("Environment variable GITHUB_TOKEN is not set.");
                Environment.Exit(1);
            }

            comboBoxWorksheet.DataSource = worksheetItems;
            comboBoxWorksheet.DisplayMember = "worksheetItemName";      // Text shown in dropdown
            comboBoxWorksheet.ValueMember = "baseImageFileNumber";    // Value you can access later
            comboBoxWorksheet.SelectedIndex = 1;

            txtWordApp.Text = SelectaUserGuide;

            _cachedImageFilesInRepo = null;
        }
        //-----------------------------------------------------------------------------------------
        private void Form1_Load(object sender, EventArgs e)
        {
            // Position at top-center of the primary screen
            int screenWidth = Screen.PrimaryScreen.WorkingArea.Width;
            int formWidth = this.Width;

            this.StartPosition = FormStartPosition.Manual;
            this.Location = new Point((screenWidth - formWidth) / 2, 5);  // X = center, Y = 0 (top)
        }
        //-----------------------------------------------------------------------------------------
        private async void btnUploadImages_Click(object sender, EventArgs e)
        {
            Stop = false;

            if (!ValidateInputSettings())
                return;

            string wordAppPath = txtWordApp.Text.Trim();
            string excelAppPath = txtExcelApp.Text.Trim();
            string sheetName = comboBoxWorksheet.Text.Trim();

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
                COL_ICON = GetColumnNumberByHeaderTitle(excelPackage.Workbook, sheetName, "Icon");
                COL_COLOR = GetColumnNumberByHeaderTitle(excelPackage.Workbook, sheetName, "Color");
                COL_TRANSFORM = GetColumnNumberByHeaderTitle(excelPackage.Workbook, sheetName, "Transform");
                COL_ALT_TEXT = GetColumnNumberByHeaderTitle(excelPackage.Workbook, sheetName, "Alt Text");
                COL_MAX_HEIGHT = GetColumnNumberByHeaderTitle(excelPackage.Workbook, sheetName, "Height Max");

                // safety check - identify writeable columns
                WritableColumnsBySheet.Clear();
                WritableColumnsBySheet.Add(sheetName, new HashSet<int> { COL_ALT_TEXT, COL_ERRORS });

                // verify that all AltTextId values are sequential

                DialogResult msgResult;
                bool keepChecking = true;

                // Verify Image Names are in sequence
                for (int row = 3; row < 1000; row++)        // first row is header
                {
                    BeginInvoke((Action)(() => labelStatus.Text = worksheet.Cells[row - 1, COL_IMAGE_NAME].Text));
                    Application.DoEvents();

                    if (!string.IsNullOrEmpty(worksheet.Cells[row, COL_IMAGE_NAME].Text))
                    {
                        int row0 = ExtractImageNumber(worksheet.Cells[row - 1, COL_IMAGE_NAME].Text);
                        int row1 = ExtractImageNumber(worksheet.Cells[row, COL_IMAGE_NAME].Text);
                        if (row0 + 1 != row1)
                        {
                            msgResult = MessageBox.Show("Out of sequence at row " + row.ToString(), "Continue?", MessageBoxButtons.OKCancel);
                            if (msgResult == DialogResult.Cancel)
                                return;
                        }
                    }
                    else
                    {
                        msgResult = MessageBox.Show("Image names are sequential. Last Image Name = " + worksheet.Cells[row-1, COL_IMAGE_NAME].Text, "Continue?", MessageBoxButtons.OKCancel);
                        if (msgResult == DialogResult.Cancel)
                            return;
                        else
                            break;
                    }
                    BeginInvoke((Action)(() => labelStatus.Text = worksheet.Cells[row - 1, COL_IMAGE_NAME].Text));
                    Application.DoEvents();
                }

                int baseImageValue = (int)comboBoxWorksheet.SelectedValue;

                for (int row = 2; row < 999; row++)        // first row is header
                {
                    Application.DoEvents();

                    if (Stop)
                    {
                        Application.DoEvents();
                        excelPackage.Save();     // background cell colors
                        MessageBox.Show("Stopped");
                        Application.DoEvents();
                        return;
                    }

                    SafeSetCellValue(worksheet, row, COL_ERRORS, "");

                    // required. stop if not found
                    string sourceImageName = worksheet.Cells[row, COL_IMAGE_NAME].Text;
                    if (string.IsNullOrWhiteSpace(sourceImageName))
                    {
                        // ensure there are no varants with this image number that may have formerly been images in git repo
                        string delImageNumber = (baseImageValue + row - 1).ToString();
                        BeginInvoke((Action)(() => labelStatus.Text = "Deleting " + delImageNumber));
                        await DeleteImagesByPrefixAsync(delImageNumber, string.Empty);
                        continue;
                    }
                    //// required. stop if not found
                    //string altTextId = worksheet.Cells[row, COL_IMAGE_ID].Text;
                    //if (string.IsNullOrWhiteSpace(altTextId))
                    //    break;





                    //BeginInvoke((Action)(() => labelWordImage.Text = sourceImageName));
                    Application.DoEvents();

                    string internalPath = "word/media/" + sourceImageName;
                    ZipArchiveEntry wordMediaImage = wordAppAsZip.GetEntry(internalPath);
                    if (wordMediaImage == null)
                    {
                        Console.WriteLine("Image not found: " + sourceImageName);
                        continue;
                    }

                    byte[] wordDocImageBytes;
                    using (Stream sourceStream = wordMediaImage.Open())
                    {
                        using (MemoryStream ms = new MemoryStream())
                        {
                            sourceStream.CopyTo(ms);
                            wordDocImageBytes = ms.ToArray();
                        }
                    }
                    //int wordDocImageHash =  wordDocImageBytes.Aggregate(17, (current, b) => current * 31 + b);
                    string wordDocImageSha = ComputeGitHubBlobSha(wordDocImageBytes);

                    Bitmap wordBitmap;
                    using (var ms = new MemoryStream(wordDocImageBytes))
                    {
                        wordBitmap = new Bitmap(ms);
                    }

                    // construct image number
                    int imageNumber = ExtractImageNumber(worksheet.Cells[row, COL_IMAGE_NAME].Text);
                    string destImageNumber = (baseImageValue + imageNumber).ToString();
                    if (string.IsNullOrWhiteSpace(destImageNumber))
                        break;

                    string destFileName = string.Empty;

                    string title = worksheet.Cells[row, COL_TITLE].Text;

                    // upload to repo only images - not icons
                    string icon = worksheet.Cells[row, COL_ICON].Text;

                    if (false == string.IsNullOrEmpty(icon))
                    {
                        // this is an icon - no image in repo required.
                        string colour = worksheet.Cells[row, COL_COLOR].Text;
                        string transform = worksheet.Cells[row, COL_TRANSFORM].Text;

                        // record the Alt Text to be written (by another process) to the User Guide
                        string altTextValue = GetAltText(icon, title, colour, transform, destImageNumber);
                        SafeSetCellValue(worksheet, row, COL_ALT_TEXT, altTextValue);

                        // ensure there are no varants with this image number that may have formerly been images in git repo
                        await DeleteImagesByPrefixAsync(destImageNumber, string.Empty);
                        continue;
                    }
                    else
                    {
                        // get the file name from Excel
                        string excelFileName = worksheet.Cells[row, COL_FILENAME].Text;
                        // next row if no File Name
                        if (string.IsNullOrWhiteSpace(excelFileName))
                            continue;

                        // ensures no duplicate file names across User Guides
                        destFileName = destImageNumber + "-" + excelFileName;
                    }

                    string json = await GetFileRepoAsync(destFileName).ConfigureAwait(false);
                    if (json == null)
                    {
                        // the exact filename does NOT exist in git repo
                        // record in listbox
                        SafeSetCellValue(worksheet, row, COL_ERRORS, "New");

                        BeginInvoke((Action)(() =>
                        {
                            listBoxCollisions.Items.Add(destFileName);
                            picBoxPanel.BackColor = Color.Red;
                            Application.DoEvents();
                            //MessageBox.Show("MisMatch: " + destFileName);
                        }));

                        // ensure there are no varants with this image number
                        await DeleteImagesByPrefixAsync(destImageNumber, string.Empty);

                        // push the image to the repo
                        bool push_result = PutFileRepoAsync(destFileName, string.Empty, wordDocImageBytes).GetAwaiter().GetResult();
                        if (!push_result)
                        {
                            Console.WriteLine("Upload failed: " + destFileName);
                            MessageBox.Show("Upload failed: " + destFileName);
                        }

                        json = await GetFileRepoAsync(destFileName).ConfigureAwait(false);

                        GitHubFileInfo fileInfo = JsonConvert.DeserializeObject<GitHubFileInfo>(json);

                        // lookup the existing image
                        string html_url = (string)fileInfo.html_url;

                        string max_height = worksheet.Cells[row, COL_MAX_HEIGHT].Text;

                        // record the Alt Text to be written (by another process) to the User Guide
                        string altTextValue = GetAltText(html_url, title, max_height, destImageNumber);
                        SafeSetCellValue(worksheet, row, COL_ALT_TEXT, altTextValue);

                        // show the images in the repo
                        BeginInvoke((Action)(() =>
                        {
                            labelFileName.Text = destFileName;
                            labelWordImage.Text = sourceImageName;
                            picBoxWordImage.Image = (Image)wordBitmap.Clone();
                            picBoxGitRepoImage.Image = null;
                            picBoxPanel.BackColor = SystemColors.Control;
                        }));
                    }
                    else
                    {
                        // the exact filename DOES exist in git repo

                        GitHubFileInfo fileInfo = JsonConvert.DeserializeObject<GitHubFileInfo>(json);

                        // lookup the existing image
                        string html_url = (string)fileInfo.html_url;

                        string max_height = worksheet.Cells[row, COL_MAX_HEIGHT].Text;

                        // record the Alt Text to be written (by another process) to the User Guide
                        string altTextValue = GetAltText(html_url, title, max_height, destImageNumber);
                        SafeSetCellValue(worksheet, row, COL_ALT_TEXT, altTextValue);

                        string base64 = (string)fileInfo.content;
                        base64 = base64.Replace("\n", "").Replace("\r", ""); // GitHub adds newlines to base64 output

                        byte[] gitRepoImageBytes = Convert.FromBase64String(base64);

                        Bitmap gitBitmap;
                        using (var ms = new MemoryStream(gitRepoImageBytes))
                        {
                            gitBitmap = new Bitmap(ms);
                        }
                        // show the images in the repo
                        BeginInvoke((Action)(() =>
                        {
                            labelFileName.Text = destFileName;
                            labelWordImage.Text = sourceImageName;
                            picBoxWordImage.Image = (Image)wordBitmap.Clone();
                            picBoxGitRepoImage.Image = (Image)gitBitmap.Clone();
                            picBoxPanel.BackColor = Color.PaleGreen;
                        }));
                        // ensure there are no varants with this filename's image number
                        await DeleteImagesByPrefixAsync(destImageNumber, destFileName);

                        // upload the image if the sha(s) don't match
                        if (fileInfo.sha != wordDocImageSha)
                        {
                            //worksheet.Cells[row, 3].Value = "Error";
                            SafeSetCellValue(worksheet, row, COL_ERRORS, "SHA mismatch");
                            BeginInvoke((Action)(() =>
                            {
                                listBoxCollisions.Items.Add(destFileName);
                                picBoxPanel.BackColor = Color.Red;
                                Application.DoEvents();
                                //MessageBox.Show("MisMatch: " + destFileName);
                            }));
                            // push the image to the repo
                            bool push_result = PutFileRepoAsync(destFileName, fileInfo.sha, wordDocImageBytes).GetAwaiter().GetResult();
                            if (!push_result)
                            {
                                Console.WriteLine("Upload failed: " + destFileName);
                                MessageBox.Show("Upload failed: " + destFileName);
                            }
                        }
                    }
                    Application.DoEvents();
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
        private async Task<List<(string name, string sha)>> ListImageFilesInRepoAsync()
        {
            if (_cachedImageFilesInRepo != null)
                return _cachedImageFilesInRepo;

            string url = $"https://api.github.com/repos/Gradescan/images/contents";

            var files = new List<(string, string)>();

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("UploaderApp", "1.0"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("token", GitHubToken);

                try
                {
                    HttpResponseMessage response = await client.GetAsync(url).ConfigureAwait(false);
                    string json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    if (!response.IsSuccessStatusCode)
                        return files;

                    dynamic items = JsonConvert.DeserializeObject(json);

                    foreach (var item in items)
                    {
                        string name = item.name;
                        string sha = item.sha;

                        if (name.EndsWith(".png"))
                            files.Add((name, sha));
                    }

                    _cachedImageFilesInRepo = files;
                    return files;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("List error: " + ex.Message);
                    return files;
                }
            }
        }
        //-----------------------------------------------------------------------------------------
        private async Task DeleteImagesByPrefixAsync(string destImageNumber, string exceptFileName)
        {
            var files = await ListImageFilesInRepoAsync();
            string fileName = $"{exceptFileName}.png";

            foreach (var (name, sha) in files)
            {
                if (name.StartsWith(destImageNumber + "-")
                 && name != fileName)
                {
                    bool deleted = await DeleteFileFromRepoAsync(name, sha);
                    if (deleted)
                        Console.WriteLine($"Deleted: {name}");
                    else
                        Console.WriteLine($"Failed to delete: {name}");

                    Application.DoEvents();
                }
            }
        }
        //-----------------------------------------------------------------------------------------
        private async Task<bool> DeleteFileFromRepoAsync(string fileName, string sha)
        {
            string url = $"https://api.github.com/repos/Gradescan/images/contents/{Uri.EscapeDataString(fileName)}";

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("UploaderApp", "1.0"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("token", GitHubToken);

                var deleteBody = new
                {
                    message = $"Delete {fileName}",
                    sha = sha
                };

                string json = JsonConvert.SerializeObject(deleteBody);
                var content = new StringContent(json, Encoding.UTF8, "application/json");

                try
                {
                    HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Delete, url) { Content = content };
                    HttpResponseMessage response = await client.SendAsync(request).ConfigureAwait(false);

                    return response.IsSuccessStatusCode;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Delete error: " + ex.Message);
                    return false;
                }
            }
        }
        //-----------------------------------------------------------------------------------------
        public static string ComputeGitHubBlobSha(byte[] contentBytes)
        {
            string header = $"blob {contentBytes.Length}\0";
            byte[] headerBytes = Encoding.UTF8.GetBytes(header);

            byte[] blob = new byte[headerBytes.Length + contentBytes.Length];
            Buffer.BlockCopy(headerBytes, 0, blob, 0, headerBytes.Length);
            Buffer.BlockCopy(contentBytes, 0, blob, headerBytes.Length, contentBytes.Length);

            using (SHA1 sha1 = SHA1.Create())
            {
                byte[] hash = sha1.ComputeHash(blob);
                string sha = BitConverter.ToString(hash).Replace("-", "").ToLowerInvariant();
                return sha;
            }
        }
        //-----------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------
        // Add the following method to Form1.cs
        private void btnCreateTxtFile_Click(object sender, EventArgs e)
        {
            if (!ValidateInputSettings())
                return;

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                string wordAppPath = txtWordApp.Text.Trim();
                string sheetName = comboBoxWorksheet.Text.Trim();
                string docNameNoExt = Path.GetFileNameWithoutExtension(wordAppPath);
                string docDir = Path.GetDirectoryName(wordAppPath);

                // Load Excel map
                string excelPath = txtExcelApp.Text.Trim();
                var dict = LoadAltTextFromExcel_Map(excelPath, sheetName);

                wordApp = new Word.Application();
                doc = wordApp.Documents.Open(wordAppPath);

                StringBuilder output = new StringBuilder();
                StringBuilder pageBuffer = new StringBuilder();

                int lastPage = -1;
                int previousPage = 0;

                foreach (Word.Paragraph para in doc.Paragraphs)
                {
                    Word.Range rng = para.Range;
                    int currentPage = rng.get_Information(Word.WdInformation.wdActiveEndPageNumber);

                    if (currentPage != lastPage && lastPage != -1)
                    {
                        pageBuffer.AppendLine($"<p>You can read the User Guide page <a href=\"{{ClientRootAddress}}/assets/docs/{EncodeURLComponent(docNameNoExt)}.pdf#page={lastPage}\" target=\"_blank\" rel=\"noopener noreferrer\">here</a>.</p>");
                        output.Append(pageBuffer);
                        pageBuffer.Clear();
                    }

                    // Replace alt text in Shapes on this paragraph
                    foreach (Word.Shape shp in doc.Shapes)
                    {
                        if (shp.Anchor.Start >= rng.Start && shp.Anchor.Start < rng.End)
                        {
                            string altText = shp.AlternativeText.Trim();
                            Match match = Regex.Match(altText, "\\[(\\d{4})\\]");
                            if (match.Success && dict.ContainsKey(match.Groups[1].Value))
                            {
                                shp.AlternativeText = dict[match.Groups[1].Value];
                            }
                        }
                    }

                    string line = "";
                    for (int i = 1; i <= rng.Words.Count; i++)
                    {
                        Word.Range word = rng.Words[i];

                        if (word.InlineShapes.Count > 0)
                        {
                            foreach (Word.InlineShape ils in word.InlineShapes)
                            {
                                if (ils.Type == Word.WdInlineShapeType.wdInlineShapePicture)
                                {
                                    string altText = ils.AlternativeText.Trim();
                                    Match match = Regex.Match(altText, "\\[(\\d{4})\\]");
                                    if (match.Success && dict.ContainsKey(match.Groups[1].Value))
                                    {
                                        ils.AlternativeText = dict[match.Groups[1].Value];
                                        altText = dict[match.Groups[1].Value];
                                    }
                                    altText = Regex.Replace(altText, "\\[(\\d{4})\\]", "");
                                    if (Regex.IsMatch(altText, "src=\\\"[^\\\"]+\\\""))
                                        line += altText + Environment.NewLine;
                                    else if (altText.Contains("<span"))
                                        line += altText + " ";
                                }
                            }
                        }
                        else
                        {
                            line += word.Text;
                        }
                    }

                    pageBuffer.Append(line);
                    lastPage = currentPage;

                    if (previousPage != currentPage)
                    {
                        Console.Write(previousPage == 0 ? $"Page {currentPage}" : $",{currentPage}");
                        previousPage = currentPage;
                    }
                }

                if (pageBuffer.Length > 0)
                {
                    pageBuffer.AppendLine($"<p>You can read the User Guide page <a href=\"{{ClientRootAddress}}/assets/docs/{EncodeURLComponent(docNameNoExt)}.pdf#page={lastPage}\" target=\"_blank\" rel=\"noopener noreferrer\">here</a>.</p>");
                    output.Append(pageBuffer);
                }

                string txtFilePath = Path.Combine(docDir, docNameNoExt + ".txt");
                File.WriteAllText(txtFilePath, SanitizeControlChars(output.ToString()).Replace("\"", "\"\""));

                MessageBox.Show("Export complete!\nText file: " + txtFilePath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                if (doc != null) { doc.Close(false); Marshal.ReleaseComObject(doc); }
                if (wordApp != null) { wordApp.Quit(); Marshal.ReleaseComObject(wordApp); }
                GC.Collect(); GC.WaitForPendingFinalizers();
            }
        }
        //-----------------------------------------------------------------------------------------
        private Dictionary<string, string> LoadAltTextFromExcel_Map(string excelPath, string sheetName)
        {
            var dict = new Dictionary<string, string>();

            using (var package = new ExcelPackage(new FileInfo(excelPath)))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[sheetName];

                for (int row = 2; row <= worksheet.Dimension.End.Row; row++)
                {
                    string key = worksheet.Cells[row, COL_IMAGE_NAME].Text; // worksheet.Cells[row, colId].Text.Trim();

                    string title = worksheet.Cells[row, COL_TITLE].Text;

                    // upload to repo only images - not icons
                    string icon = worksheet.Cells[row, COL_ICON].Text;

                    if (false == string.IsNullOrEmpty(icon))
                    {
                        // this is an icon - no image in repo required.
                        string colour = worksheet.Cells[row, COL_COLOR].Text;
                        string transform = worksheet.Cells[row, COL_TRANSFORM].Text;
                        string altTextId = ExtractImageNumber(worksheet.Cells[row, COL_IMAGE_NAME].Text).ToString();
                        // record the Alt Text to be written (by another process) to the User Guide
                        string altTextValue = GetAltText(icon, title, colour, transform, altTextId);
                        dict[key] = altTextValue;
                   //     SafeSetCellValue(worksheet, row, COL_ALT_TEXT, altTextValue);
                    }
                    //string val = worksheet.Cells[row, colText].Text.Trim();
                    //if (!string.IsNullOrEmpty(key)) dict[key] = val;
                }
            }
            return dict;
        }
        //-----------------------------------------------------------------------------------------
        private string EncodeURLComponent(string s)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in s)
            {
                if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
                    sb.Append(c);
                else if (c == ' ')
                    sb.Append("%20");
                else
                    sb.Append("%" + ((int)c).ToString("X2"));
            }
            return sb.ToString();
        }
        //-----------------------------------------------------------------------------------------
        private string SanitizeControlChars(string text)
        {
            StringBuilder sb = new StringBuilder();
            foreach (char c in text)
            {
                if (c == 9 || c == 10 || c == 13 || c >= 32)
                    sb.Append(c);
            }
            return sb.ToString();
        }
        //-----------------------------------------------------------------------------------------
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
                    string resultStr = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    GitReadStatus result = JsonConvert.DeserializeObject<GitReadStatus>(resultStr);

                    System.Diagnostics.Debug.WriteLine("IsSuccessStatusCode: " + response.IsSuccessStatusCode);
                    System.Diagnostics.Debug.WriteLine("ReasonPhrase: " + response.ReasonPhrase);
                    System.Diagnostics.Debug.WriteLine("result.message: " + result.message);

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
        private void buttonStop_Click(object sender, EventArgs e)
        {
            Stop = true;
        }
        //-----------------------------------------------------------------------------------------
        //<img src="https://github.com/Gradescan/images/blob/main/question-edit-answers-ordered-222.png?raw=true" style="max-height: 200px; width: auto;" alt="Question Edit">

        private string GetAltText(string html_url, string title, string max_height, string altTextId)
        {
            string _maxheight = (string.IsNullOrEmpty(max_height)
                              ? $@"style=""max-width: auto; width: auto; "" "
                              : $@"style=""max-width: auto; width: auto; max-height: {max_height}px; "" ");
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
            if (!ValidateInputSettings())
                return;

            // Get the starting number from comboBoxWorksheet
            int startingValue = (int)comboBoxWorksheet.SelectedValue + 1;
            int counter = startingValue;
            int largestNumber = counter;

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                string wordAppPath = txtWordApp.Text.Trim();
                string sheetName = comboBoxWorksheet.Text.Trim();

                wordApp = new Word.Application();
                doc = wordApp.Documents.Open(wordAppPath);

                // match legacy 3-digit numbers or current standar 4-digit numbers.
                Regex pattern = new Regex(@"\[(\d{3,4})\]");


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
                        shape.AlternativeText = altText + "[" + counter.ToString("D4") + "]";
                        counter++;
                        largestNumber = counter;
                    }
                    // show the status
                    BeginInvoke((Action)(() => labelStatus.Text = counter.ToString()));
                    Application.DoEvents();
                }

                foreach (Word.Shape shape in doc.Shapes)
                {
                    string altText = shape.AlternativeText;
                    if (!pattern.IsMatch(altText))
                    {
                        string newText = altText + "\n[" + counter.ToString("D4") + "]";
                        shape.AlternativeText = newText;
                        counter++;
                    }
                    // show the status
                    BeginInvoke((Action)(() => labelStatus.Text = counter.ToString()));
                    Application.DoEvents();
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
            if (!ValidateInputSettings())
                return;

            Word.Application wordApp = null;
            Word.Document doc = null;

            try
            {
                string wordAppPath = txtWordApp.Text.Trim();
                wordApp = new Word.Application();
                doc = wordApp.Documents.Open(wordAppPath);

                int counter = 0;
                int convertedCount = 0;

                // Clear InlineShapes Alt Text
                foreach (Word.InlineShape shape in doc.InlineShapes)
                {
                    if (!string.IsNullOrEmpty(shape.AlternativeText))
                    {
                        shape.AlternativeText = string.Empty;
                        counter++;
                    }
                    // show the status
                    BeginInvoke((Action)(() => labelStatus.Text = counter.ToString()));
                    Application.DoEvents();
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
                        counter++;
                    }
                    // show the status
                    BeginInvoke((Action)(() => labelStatus.Text = counter.ToString()));
                    Application.DoEvents();
                }

                doc.Save();
                MessageBox.Show($"Cleared AltText from {counter} shapes.  Converted {convertedCount} shapes.");
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
        private void btnVerify_Click(object sender, EventArgs e)
        {
            if (!ValidateInputSettings())
                return;

        }
        //-----------------------------------------------------------------------------------------
        private int ExtractImageNumber(string input)
        {
            var match = Regex.Match(input, @"^image(\d{1,3})\.png$", RegexOptions.IgnoreCase);
            if (match.Success && int.TryParse(match.Groups[1].Value, out int result))
            {
                return result;
            }
            throw new ArgumentException("Input string is not in the expected format: imageDDD.png");
        }
        //-----------------------------------------------------------------------------------------
        private bool ValidateInputSettings()
        {
            string wordAppPath = txtWordApp.Text.Trim();
            string excelAppPath = txtExcelApp.Text.Trim();
            string sheetName = comboBoxWorksheet.Text.Trim();

            if (IsFileOpen(excelAppPath))
            {
                MessageBox.Show("The Excel file is currently open. Please save and close the file.");
                return false;
            }

            if (SelectaUserGuide == wordAppPath)
            {
                MessageBox.Show(SelectaUserGuide);
                return false;
            }

            if (IsFileOpen(wordAppPath))
            {
                MessageBox.Show("The User Guide file is currently open. Please save and close the file.");
                return false;
            }

            if (!wordAppPath.Contains(sheetName))
            {
                MessageBox.Show("Word File and Worksheet Name mismatch");
                return false;
            }

            return true;
        }
        //-----------------------------------------------------------------------------------------
        public static bool IsFileOpen(string filePath)
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
        public class WorksheetItem
        {
            public string  worksheetItemName { get; set; }
            public int     baseImageFileNumber { get; set; }

            public WorksheetItem(string name, int imgnum)
            {
                worksheetItemName = name;
                baseImageFileNumber = imgnum;
            }
            public override string ToString()
            {
                return worksheetItemName;
            }
        }
    }
    //-----------------------------------------------------------------------------------------
}
