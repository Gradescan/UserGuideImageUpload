using System;
using System.IO;
using System.IO.Compression;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows.Forms;
using OfficeOpenXml;

namespace ExcelWordImageUploader
{
    public partial class Form1 : Form
    {
        //-----------------------------------------------------------------------------------------
        private const string DefaultGitHubFolder = "Gradescan Professional User Guide";

        private string GitHubToken;
        //-----------------------------------------------------------------------------------------
        public Form1()
        {
            InitializeComponent();

#pragma warning disable CS0618
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
#pragma warning restore CS0618

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
            string repoPath = txtRepo.Text.Trim();
            string excelPath = txtExcel.Text.Trim();
            string wordPath = txtWord.Text.Trim();
            string sheetName = txtSheet.Text.Trim();

            if (!repoPath.Contains("/"))
            {
                MessageBox.Show("Repo must be in format: owner/repo (e.g. Gradescan/images)");
                return;
            }

            string[] parts = repoPath.Split('/');
            string gitUsername = parts[0];
            string gitRepoName = parts[1];

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
                        bool result = UploadToGitHub(
                            gitUsername, gitRepoName, destFileName, buffer).GetAwaiter().GetResult();

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

        private async Task<bool> UploadToGitHub(string username, string repo, string fileName, byte[] content)
        {
            string path = Uri.EscapeUriString(DefaultGitHubFolder + "/" + fileName);
            string url = $"https://api.github.com/repos/{username}/{repo}/contents/{path}";

            string base64 = Convert.ToBase64String(content);
            var payload = new
            {
                message = "Auto upload " + fileName,
                content = base64
            };

            string json = JsonSerializer.Serialize(payload);
            HttpClient client = new HttpClient();

            client.DefaultRequestHeaders.UserAgent.Add(new ProductInfoHeaderValue("UploaderApp", "1.0"));
            client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("token", GitHubToken);

            HttpContent httpContent = new StringContent(json, Encoding.UTF8, "application/json");
            HttpResponseMessage response = await client.PutAsync(url, httpContent);

            return response.IsSuccessStatusCode;
        }
        //-----------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------
        //-----------------------------------------------------------------------------------------
    }
    //-----------------------------------------------------------------------------------------
}
