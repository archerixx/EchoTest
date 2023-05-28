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

            return Challenge(properties, GoogleDefaults.AuthenticationScheme);
        }

        [HttpGet("parse-token")]
        public async Task<IActionResult> GetTokenAsync()
        {
            var loggedInUser = await HttpContext.AuthenticateAsync();
            if (loggedInUser.Succeeded)
            {
                // fetch googletoken
                //HttpContext.Response.Cookies.Append("Token", "Cookie Value");
            }

            return LocalRedirect("/");
        }
    }
}
