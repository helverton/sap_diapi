using Owin;
using Hangfire;
using Microsoft.Owin;
using System.Globalization;
using HelvertonSantos.Controllers;

[assembly: OwinStartup(typeof(HelvertonSantos.Startup))]

namespace HelvertonSantos
{
    public class Startup
    {
        public void Configuration(IAppBuilder app)
        {
            string conf = $"Server={System.Configuration.ConfigurationManager.AppSettings["Server"]}; " +
                          $"Database ={ System.Configuration.ConfigurationManager.AppSettings["Database"]}; " +
                          $"User Id = { System.Configuration.ConfigurationManager.AppSettings["User"] }; " +
                          $"Password ={ System.Configuration.ConfigurationManager.AppSettings["Password"]}; ";
            GlobalConfiguration.Configuration
                .UseSqlServerStorage(conf);

            CultureInfo.DefaultThreadCurrentCulture = new CultureInfo("en-US");

            app.UseHangfireDashboard();
            app.UseHangfireServer();

            HangFireController.Start();
        }
    }
}
