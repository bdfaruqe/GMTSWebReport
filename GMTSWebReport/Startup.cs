using Microsoft.Owin;
using Owin;

[assembly: OwinStartupAttribute(typeof(GMTSWebReport.Startup))]
namespace GMTSWebReport
{
    public partial class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            ConfigureAuth(app);
        }
    }
}
