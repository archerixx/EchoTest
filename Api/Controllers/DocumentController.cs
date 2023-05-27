using Api.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace Api.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
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
                var a = HttpContext.Request.Headers.FirstOrDefault(v => v.Key == "Authorization");

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
        public IActionResult GetGoogleSheet()
        {
            try
            {

                var a = HttpContext.Request.Headers.FirstOrDefault(v => v.Key == "Authorization");
                return Ok(documentService.GetGoogleSheet());
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

                var a = HttpContext.Request.Headers.FirstOrDefault(v => v.Key == "Authorization");
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
