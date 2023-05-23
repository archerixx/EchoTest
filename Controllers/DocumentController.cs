using EchoTest.Interfaces;
using Microsoft.AspNetCore.Mvc;

namespace EchoTest.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class DocumentController : Controller
    {
        private readonly IDocumentService documentService;

        public DocumentController(IDocumentService documentService)
        {
            this.documentService = documentService;
        }

        [HttpGet("GenerateGoogleDoc")]
        public IActionResult GenerateGoogleDoc(string googleSheetUrl = "18U4J_Zzhanm-oangtg31VMcIF9ZuuGd5dYw_7vkNC2g", string destination = null)
        {
            try
            {
                documentService.GenerateAndUploadGoogleDocument(googleSheetUrl, destination);
                return Ok();
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
    }
}
