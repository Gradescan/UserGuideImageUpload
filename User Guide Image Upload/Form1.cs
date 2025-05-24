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

        private void btnRun_Click(object sender, EventArgs e)
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

                    using (Stream sourceStream = entry.Open())
                    using (MemoryStream ms = new MemoryStream())
                    {
                        sourceStream.CopyTo(ms);
                        byte[] buffer = ms.ToArray();
                        bool result = UploadToGitHub(repoFolder, destFileName, buffer).GetAwaiter().GetResult();

                        if (!result)
                            Console.WriteLine("Upload failed: " + destFileName);
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

                Bitmap bitmap;
                using (var ms = new MemoryStream(imageBytes))
                {
                    bitmap = new Bitmap(ms);
                }

                string base64 = Convert.ToBase64String(imageBytes);

                var payload = new
                {
                    message = "Auto upload " + fileName,
                    content = base64,
                    committer = new
                    {
                        name = "UploaderBot",
                        email = "uploader@gradescan.org"
                    },
                    author = new
                    {
                        name = "UploaderBot",
                        email = "uploader@gradescan.org"
                    }
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
        //-----------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------
    }
    //-----------------------------------------------------------------------------------------
}
