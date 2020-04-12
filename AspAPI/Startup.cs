using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(AspAPI.Startup))]
namespace AspAPI
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
