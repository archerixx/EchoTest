using EchoTest.Helper;
using EchoTest.Interfaces;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Docs.v1;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Util.Store;
using Word = Microsoft.Office.Interop.Word;

namespace EchoTest.Services
{
    public class DocumentService : IDocumentService
    {
        private readonly IConfiguration configuration;

        public DocumentService(IConfiguration configuration)
        {
            this.configuration = configuration;
        }

        public void GenerateAndUploadGoogleDocument(string spreadsheetId, string destination)
        {

            Spreadsheet spreadsheet = GetSpreadsheet(spreadsheetId);
            var currentPath = Directory.GetCurrentDirectory();
            var filePath = Path.Combine(currentPath, "Assets\\fileformat.docx");
            GenerateWordHelper.GenerateWordDocumentFromSpreadsheet(spreadsheet, filePath);
            // update fields (TOC, Tables)
            UpdateFieldsInWordDocument(filePath);
            // upload to GoogleDocs
            UploadToGoogleDocs(filePath);
        }

        private Spreadsheet GetSpreadsheet(string spreadsheetId)
        {
            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                ApiKey = configuration.GetSection("GoogleAPI:ApiKey").Value,
                ApplicationName = "EchoTest",
            });

            // TODO - check what ranges are ???
            // The ranges to retrieve from the spreadsheet.
            // List<string> ranges = new List<string>();

            SpreadsheetsResource.GetRequest request = sheetsService.Spreadsheets.Get(spreadsheetId);
            //request.Ranges = ranges;
            request.IncludeGridData = true; // if grid data should be returned.

            return request.Execute();
        }

        private void UpdateFieldsInWordDocument(string filePath)
        {
            // Create an instance of Word application
            var wordApp = new Word.Application();

            // Disable alerts
            wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsNone;

            // Open the Word document
            var wordDoc = wordApp.Documents.Open(filePath);

            // Update fields
            wordDoc.Fields.Update();

            // Save and close the document
            wordDoc.Save();
            wordDoc.Close();

            // Enable alerts
            wordApp.DisplayAlerts = Word.WdAlertLevel.wdAlertsAll;

            // Quit Word application
            wordApp.Quit();
        }

        private static UserCredential GetGoogleApiCredential()
        {
            using (var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read))
            {
                var currDir = Directory.GetCurrentDirectory();
                string credPath = Path.Join(currDir, "test.json");

                return GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    new string[] { DocsService.Scope.Documents, DocsService.Scope.Drive, DocsService.Scope.DriveFile },
                    "testnijob@gmail.com",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)
                    ).Result;
            }
        }

        private static void UploadToGoogleDocs(string filePath)
        {
            #region Driver test

            DriveService driveService = new DriveService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetGoogleApiCredential(),
                ApplicationName = "EchoTest",
            });

            Google.Apis.Drive.v3.Data.File body = new();
            body.Name = Path.GetFileName(filePath);
            //body.Capabilities = new CapabilitiesData { CanEdit = true };
            // Set what the end mime type will be
            body.MimeType = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
            byte[] byteArray = File.ReadAllBytes(filePath);

            MemoryStream streamFile = new MemoryStream(byteArray);
            FilesResource.CreateMediaUpload request =
                driveService.Files.Create(
                    body,
                    streamFile,
                    "application/vnd.openxmlformats-officedocument.wordprocessingml.document"); // the upload format
            request.Upload();
            #endregion

        }
    }
}
