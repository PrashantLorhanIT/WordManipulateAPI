using System;
using System.Threading.Tasks;
using Microsoft.Owin;
using Owin;
using Microsoft.Owin.Security.Jwt;
using Microsoft.Owin.Security;
using Microsoft.IdentityModel.Tokens;
using System.Text;
using System.Configuration;
[assembly: OwinStartup(typeof(WordManipulateAPI.Startup))]

namespace WordManipulateAPI
{
    public class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            // For more information on how to configure your application, visit https://go.microsoft.com/fwlink/?LinkID=316888

       
            var _secret = ConfigurationManager.AppSettings["secret"];
            var _expDate = ConfigurationManager.AppSettings["expirationInMinutes"];  
            var _audience = ConfigurationManager.AppSettings["audience"]; 
            var _issuer = ConfigurationManager.AppSettings["issuer"];  

            app.UseJwtBearerAuthentication(
                new JwtBearerAuthenticationOptions
                {
                    AuthenticationMode = AuthenticationMode.Active,
                    TokenValidationParameters = new TokenValidationParameters()
                    {
                        ValidateIssuer = true,
                        ValidateAudience = true,
                        ValidateIssuerSigningKey = true,
                        ValidIssuer = _issuer, //some string, normally web url,  
                        ValidAudience = _audience,
                        IssuerSigningKey = new SymmetricSecurityKey(Encoding.UTF8.GetBytes(_secret)),
                        
                    }
                });
        }
    }
}
