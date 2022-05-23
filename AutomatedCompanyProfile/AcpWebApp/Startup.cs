using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(AcpWebApp.Startup))]
namespace AcpWebApp
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
