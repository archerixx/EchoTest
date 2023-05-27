using Google.Apis.Auth.OAuth2;
using Google.Apis.Docs.v1;
using Google.Apis.Drive.v3;
using Google.Apis.Drive.v3.Data;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Google.Apis.Sheets.v4.Data;
using Google.Apis.Upload;
using Google.Apis.Util.Store;
using File = Google.Apis.Drive.v3.Data.File;

namespace Api.HelperMethods
{
    public class GoogleApiHelper
    {
        private static UserCredential GetGoogleApiCredential()
        {
            using (var stream = new FileStream("client_web.json", FileMode.Open, FileAccess.Read))
            {
                var currDir = Directory.GetCurrentDirectory();
                string credPath = Path.Join(currDir);

                return GoogleWebAuthorizationBroker.AuthorizeAsync(
                    GoogleClientSecrets.Load(stream).Secrets,
                    new string[] { SheetsService.Scope.Spreadsheets, DocsService.Scope.Documents, DriveService.Scope.Drive, DocsService.Scope.DriveFile },
                    "gmail",
                    CancellationToken.None,
                    new FileDataStore(credPath, true)
                    ).Result;
            }
        }

        public static void Logout()
        {
            try
            {
                // delete cashed credentials
                var currDir = Directory.GetCurrentDirectory();
                string credPath = Path.Join(currDir, "Google.Apis.Auth.OAuth2.Responses.TokenResponse-gmail");
                System.IO.File.Delete(credPath);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }

            var credentials = GetGoogleApiCredential();
            var token = credentials.Token.AccessToken;
            credentials.Flow.RevokeTokenAsync(credentials.UserId, token, CancellationToken.None);
        }

        public static IEnumerable<File> GetSpreadsheet()
        {
            DriveService driveService = new DriveService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetGoogleApiCredential(),
                ApplicationName = "EchoTest",
            });

            // Set what the end mime type will be
            var mimeType = "application/vnd.google-apps.spreadsheet";
            var driveFilesRequest = driveService.Files.List();
            var driveFiles = driveFilesRequest.Execute();
            return driveFiles.Files.Where(f => f.MimeType == mimeType && f.Trashed != true);
        }

        public static Spreadsheet GetSpreadsheet(string spreadsheetId)
        {
            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                //ApiKey = "",//configuration.GetSection("GoogleAPI:ApiKey").Value,
                HttpClientInitializer = GetGoogleApiCredential(),
                ApplicationName = "EchoTest",
            });

            SpreadsheetsResource.GetRequest request = sheetsService.Spreadsheets.Get(spreadsheetId);
            request.IncludeGridData = true; // if grid data should be returned.

            return request.Execute();
        }

        public static string UploadToGoogleDocs(string filePath, string fileName, string? shareWith)
        {
            DriveService driveService = new DriveService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetGoogleApiCredential(),
                ApplicationName = "EchoTest",
            });

            var body = SetupFile(fileName);
            byte[] byteArray = System.IO.File.ReadAllBytes(filePath);
            MemoryStream streamFile = new MemoryStream(byteArray);

            FilesResource.CreateMediaUpload request =
                driveService.Files.Create(
                    body,
                    streamFile,
                    body.MimeType); // the upload format

            request.Fields = "id";
            var response = request.Upload();
            if (response.Status == UploadStatus.Failed)
                return string.Empty;

            var file = request.ResponseBody;
            AddPermissions(driveService, file.Id, shareWith);
            return file.Id;
        }

        public static string UploadToGoogleDocs(MemoryStream stream, string fileName, string? shareWith)
        {
            DriveService driveService = new DriveService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetGoogleApiCredential(),
                ApplicationName = "EchoTest",
            });

            var body = SetupFile(fileName);
            FilesResource.CreateMediaUpload request =
                driveService.Files.Create(
                    body,
                    stream,
                    body.MimeType); // the upload format

            request.Fields = "id";
            var response = request.Upload();

            if (response.Status == UploadStatus.Failed)
                return string.Empty;

            var file = request.ResponseBody;
            AddPermissions(driveService, file.Id, shareWith);
            return file.Id;
        }

        private static void AddPermissions(DriveService driveService, string id, string? shareWith)
        {
            if (shareWith != null)
                foreach (var email in shareWith.Split(","))
                {
                    var permission = new Permission()
                    {
                        EmailAddress = email,
                        Type = "user",
                        Role = "reader"
                    };

                    var permissionRequest = driveService.Permissions.Create(permission, id);
                    var a = permissionRequest.Execute();
                }
        }

        private static File SetupFile(string fileName)
        {
            File body = new()
            {
                Name = $"{fileName}.docx", // required extensions, docx is standard
                MimeType = "application/vnd.google-apps.document" // google doc type
            };

            return body;
        }
    }
}
