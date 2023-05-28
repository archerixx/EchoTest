using Google.Apis.Auth.AspNetCore3;
using Microsoft.AspNetCore.Authentication;
using Microsoft.AspNetCore.Authentication.Google;
using Microsoft.AspNetCore.Mvc;

namespace Web.Controllers
{
    [Route("/[controller]")]
    [ApiController]
    public class AuthController : ControllerBase
    {
        [HttpGet("google-login")]
        public IActionResult GoogleSignIn()
        {
            var properties = new AuthenticationProperties()
            {
                RedirectUri = "/Auth/parse-token"
            };

            return Challenge(properties, GoogleOpenIdConnectDefaults.AuthenticationScheme);
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
    }
}
