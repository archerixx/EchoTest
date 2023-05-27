using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Google;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using Microsoft.Net.Http.Headers;
using System.Collections.Specialized;

namespace Web.Controllers
{
    [Route("/[controller]")]
    [ApiController]
    public class AuthController : ControllerBase
    {
        private readonly IDocumentService documentService;
        private readonly ILogger<DocumentController> logger;

        public AuthController(IDocumentService documentService, ILogger<DocumentController> logger)
        {
            this.documentService = documentService;
            this.logger = logger;
        }


        [HttpGet("google-login")]
        public IActionResult GoogleSignIn()
        {
            var properties = new AuthenticationProperties() { 
                RedirectUri = "/Auth/parse-token"
            };

            return Challenge(properties, GoogleDefaults.AuthenticationScheme);
        }

        [HttpGet("parse-token")]
        public async Task<IActionResult> GetTokenAsync()
        {
            var loggedInUser = await HttpContext.AuthenticateAsync();
            if (loggedInUser.Succeeded)
            {
                HttpContext.Response.Cookies.Append("MyCookie", loggedInUser.Properties.Items[".Token.access_token"]);
            }

            return LocalRedirect("/");
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
