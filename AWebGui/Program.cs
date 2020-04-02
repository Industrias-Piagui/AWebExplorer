using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore;
using Microsoft.AspNetCore.Hosting;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging;

namespace AWebGui
{
    public class Program
    {
        public static void Main(string[] args)
        {
            CreateWebHostBuilder(args).Build().Run();
        }

        public static IWebHostBuilder CreateWebHostBuilder(string[] args) =>
            WebHost.CreateDefaultBuilder(args)
            .ConfigureLogging((webhostContext, builder) =>
            {
                builder.AddConfiguration(webhostContext.Configuration.GetSection("Logging"))
                .AddConsole()
                .AddDebug();
            })
            .UseStartup<Startup>()
            .UseIISIntegration();
    }
}
