using Newtonsoft.Json;
using OfficeOpenXml;
using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Web;
using System.Windows.Forms;

namespace ExcelWordImageUploader
{
    public partial class Form1 : Form
    {
        //-----------------------------------------------------------------------------------------
        private string GitHubToken;
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
        }
        //-----------------------------------------------------------------------------------------
        private void Form1_Load(object sender, EventArgs e)
        {
            // Optional logic when the form loads
        }
        //-----------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------

        private async void btnRun_Click(object sender, EventArgs e)
        {
            string repoFolder = txtRepo.Text.Trim();
            string excelPath = txtExcel.Text.Trim();
            string wordPath = txtWord.Text.Trim();
            string sheetName = txtSheet.Text.Trim();

            ZipArchive zip = null;
            ExcelPackage package = null;

            try
            {
                zip = ZipFile.OpenRead(wordPath);
                package = new ExcelPackage(new FileInfo(excelPath));
                var worksheet = package.Workbook.Worksheets[sheetName];

                for (int row = 1; row < 1000; row++)
                {
                    string sourceFileName = worksheet.Cells[row, 2].Text; // Column B
                    string destFileName = worksheet.Cells[row, 4].Text;   // Column D

                    if (string.IsNullOrWhiteSpace(sourceFileName))
                        break;

                    if (string.IsNullOrWhiteSpace(destFileName))
                        continue;

                    string internalPath = "word/media/" + sourceFileName;
                    ZipArchiveEntry entry = zip.GetEntry(internalPath);
                    if (entry == null)
                    {
                        Console.WriteLine("Image not found: " + sourceFileName);
                        continue;
                    }

                    byte[] newImageBytes;
                    using (Stream sourceStream = entry.Open())
                    {
                        using (MemoryStream ms = new MemoryStream())
                        {
                            sourceStream.CopyTo(ms);
                            newImageBytes = ms.ToArray();
                        }
                    }

                    Bitmap bitmap;
                    using (var ms = new MemoryStream(newImageBytes))
                    {
                        bitmap = new Bitmap(ms);
                    }
                    //picBoxNewImage.Image = bitmap;
                    //label2.Text = destFileName;

                    string owner = "Gradescan";
                    string repo = "media";
                    string repoPath = string.Join("/",
                        repoFolder.Split('/')
                               .Select(Uri.EscapeDataString)
                               .Concat(new[] { Uri.EscapeDataString(destFileName) })
                    );


                    string repoPathRaw = $"{repoFolder}/{destFileName}";
                    string objinfo = await GetFileRepoAsync(owner, repo, repoPathRaw).ConfigureAwait(false);
                    if (objinfo != null)
                    {
                        dynamic obj = JsonConvert.DeserializeObject(objinfo);
                        string base64 = (string)obj.content;
                        base64 = base64.Replace("\n", "").Replace("\r", ""); // GitHub adds newlines to base64 output

                        byte[] oldImageBytes = Convert.FromBase64String(base64);

                        //        byte[] fileContent = await DownloadFileFromGitHubAsync("Gradescan", "media", repoPath);

                        using (Stream sourceStream = entry.Open())
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
                        picBoxOldImage.Image = bitmap;
                    }
                    else
                    {
                        picBoxOldImage.Image = null;
                    }
                    bool result = PutFileRepoAsync(repoFolder, destFileName, newImageBytes).GetAwaiter().GetResult();

                    if (!result)
                    {
                        Console.WriteLine("Upload failed: " + destFileName);
                        MessageBox.Show("Upload failed: " + destFileName);
                    }
                }

                MessageBox.Show("Upload process completed.");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Error: " + ex.Message);
            }
            finally
            {
                if (zip != null) zip.Dispose();
                if (package != null) package.Dispose();
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
        private async Task<bool> UploadToGitHub(string repoDir, string fileName, byte[] imageBytes)
        {
            try
            {
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12;

                string owner = "Gradescan";
                string repo = "media";
                string repoPath = string.Join("/",
                    repoDir.Split('/')
                           .Select(Uri.EscapeDataString)
                           .Concat(new[] { Uri.EscapeDataString(fileName) })
                );
                string url = $"https://api.github.com/repos/{owner}/{repo}/contents/{repoPath}";

                string base64 = Convert.ToBase64String(imageBytes);

                string repoPathRaw = $"{repoDir}/{fileName}";
                string objinfo = await GetFileRepoAsync(owner, repo, repoPathRaw).ConfigureAwait(false);
                if (objinfo == null)
                    return false;

                dynamic obj = JsonConvert.DeserializeObject(objinfo);
                string sha = (string)obj.sha;
                System.Diagnostics.Debug.WriteLine("SHA: " + sha);

                return true;
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("UploadToGitHub error: " + ex.Message);
                return false;
            }
        }
        //-----------------------------------------------------------------------------------------
        private async Task<bool> PutFileRepoAsync(string repoDir, string fileName, byte[] imageBytes)
        {
            try
            {
                string owner = "Gradescan";
                string repo = "media";
                string repoPath = string.Join("/",
                    repoDir.Split('/')
                           .Select(Uri.EscapeDataString)
                           .Concat(new[] { Uri.EscapeDataString(fileName) })
                );
                string url = $"https://api.github.com/repos/{owner}/{repo}/contents/{repoPath}";

                Bitmap bitmap;
                using (var ms = new MemoryStream(imageBytes))
                {
                    bitmap = new Bitmap(ms);
                }
                picBoxNewImage.Image = bitmap;
                BeginInvoke((Action)(() => label2.Text = fileName));

                string base64 = Convert.ToBase64String(imageBytes);

                string repoPathRaw = $"{repoDir}/{fileName}";
                string objinfo = await GetFileRepoAsync(owner, repo, repoPathRaw).ConfigureAwait(false);

                string sha = string.Empty;
                if (objinfo != null)
                {
                    dynamic obj = JsonConvert.DeserializeObject(objinfo);
                    sha = (string)obj.sha;
                    System.Diagnostics.Debug.WriteLine("SHA: " + sha);
                }
                var payload = new
                {
                    message = "Auto upload " + fileName,
                    content = base64,
                    sha = sha, // <- only populated if the file exists
                    committer = new { name = "UploaderBot", email = "uploader@gradescan.org" },
                    author = new { name = "UploaderBot", email = "uploader@gradescan.org" }
                };

                string json = JsonConvert.SerializeObject(payload);

                System.Diagnostics.Debug.WriteLine("URL: " + url);
                System.Diagnostics.Debug.WriteLine("Payload: " + json);

                using (var client = new HttpClient { Timeout = TimeSpan.FromSeconds(15) })
                {
                    client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("UploaderApp", "1.0"));
                    client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("token", GitHubToken);

                    HttpContent httpContent = new StringContent(json, Encoding.UTF8, "application/json");
                    HttpResponseMessage response = await client.PutAsync(url, httpContent).ConfigureAwait(false);
                    string result = await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                    return response.IsSuccessStatusCode;
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine("UploadToGitHub error: " + ex.Message);
                return false;
            }
        }
        //-----------------------------------------------------------------------------------------
        private async Task<string> GetFileRepoAsync(string owner, string repo, string path)
        {
            string url = $"https://api.github.com/repos/{owner}/{repo}/contents/{Uri.EscapeDataString(path)}";

            using (var client = new HttpClient())
            {
                client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("UploaderApp", "1.0"));
                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("token", GitHubToken);

                try
                {
                    HttpResponseMessage response = await client.GetAsync(url).ConfigureAwait(false);

                    string json = await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    System.Diagnostics.Debug.WriteLine("SHA lookup response: " + json);

                    if (!response.IsSuccessStatusCode)
                        return null;

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
        private void btnBrowseExcel_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "Select Excel File";
            dlg.Filter = "Excel Files (*.xlsx)|*.xlsx";
            dlg.InitialDirectory = Path.GetDirectoryName(txtExcel.Text);
            if (dlg.ShowDialog() == DialogResult.OK)
                txtExcel.Text = dlg.FileName;
        }
        //-----------------------------------------------------------------------------------------
        private void btnBrowseWord_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();
            dlg.Title = "Select Word File";
            dlg.Filter = "Word Files (*.docx)|*.docx";
            dlg.InitialDirectory = Path.GetDirectoryName(txtExcel.Text);
            if (dlg.ShowDialog() == DialogResult.OK)
                txtWord.Text = dlg.FileName;
        }
        //-----------------------------------------------------------------------------------------
        private void btnBrowseRepo_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog dlg = new FolderBrowserDialog();
            if (dlg.ShowDialog() == DialogResult.OK)
                txtRepo.Text = dlg.SelectedPath.Replace('\\', '/'); // optional
        }

        private void Form1_Load_1(object sender, EventArgs e)
        {

        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
        //-----------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------
    }
    //-----------------------------------------------------------------------------------------
}
