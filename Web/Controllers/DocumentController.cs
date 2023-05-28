using Domain.Dto;
using Google.Apis.Auth.AspNetCore3;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Drive.v3;
using Google.Apis.Services;
using Microsoft.AspNetCore.Mvc;
using Web.HelperMethods;
using Web.Interfaces;

namespace Web.Controllers
{
    [ApiController]
    [Route("/[controller]")]
    public class DocumentController : Controller
    {
        private readonly IDocumentService documentService;
        private readonly ILogger<DocumentController> logger;

        public DocumentController(IDocumentService documentService, ILogger<DocumentController> logger)
        {
            this.documentService = documentService;
            this.logger = logger;
        }

        [HttpGet("GenerateGoogleDoc")]
        public IActionResult GenerateGoogleDoc(string googleSheetUrl, string? shareWith = null)
        {
            try
            {
                var docuemntId = documentService.GenerateAndUploadGoogleDocument(googleSheetUrl, shareWith);
                if (!String.IsNullOrEmpty(docuemntId))
                    return Ok(docuemntId);
                return BadRequest("Failed to generate document, please try again later");
            }
            catch (Exception ex)
            {
                logger.LogError($"FAILED - {Request.Path} - Error message: {ex}");
                return BadRequest("Something went wrong with document upload, please contact system administrator");
            }
        }

        [HttpGet("GoogleSheet")]
        [GoogleScopedAuthorize(DriveService.ScopeConstants.Drive)]
        public async Task<IActionResult> GetGoogleSheet([FromServices] IGoogleAuthProvider auth)
        //[HttpGet("GoogleSheet")]
        //public IActionResult GetGoogleSheet()
        {
            try
            {
                GoogleCredential cred = await auth.GetCredentialAsync();
                DriveService driveService = new DriveService(new BaseClientService.Initializer
                {
                    HttpClientInitializer = cred
                });

                // Set what the end mime type will be
                var mimeType = "application/vnd.google-apps.spreadsheet";
                var driveFilesRequest = driveService.Files.List();
                var driveFiles = driveFilesRequest.Execute();
                var files = driveFiles.Files.Where(f => f.MimeType == mimeType && f.Trashed != true);
                var sheetFiles = GoogleApiHelper.GetSpreadsheet();
                return Ok(sheetFiles.Select(sf => new SpreadSheetDto
                {
                    Id = sf.Id,
                    Name = sf.Name
                }));
                //return Ok(documentService.GetGoogleSheet());
            }
            catch (Exception ex)
            {
                logger.LogError($"FAILED - {Request.Path} - Error message: {ex}");
                return BadRequest("Something went wrong while retrieving documents, please contact system administrator");
            }
        }

        [HttpGet("Logout")]
        public IActionResult LogoutFromGoogleAPI()
        {
            try
            {
                documentService.LogoutFromGoogleAPI();
                return Ok();
            }
            catch (Exception ex)
            {
                logger.LogError($"FAILED - {Request.Path} - Error message: {ex}");
                return BadRequest("Something went wrong, please contact system administrator");
            }
        }
    }
}
