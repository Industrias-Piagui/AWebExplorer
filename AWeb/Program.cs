using System;
using System.Collections.Generic;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace AWeb
{
    static class Program
    {
        /// <summary>
        /// Punto de entrada principal para la aplicación.
        /// </summary>
        static void Main(string[] args)
        {
            if (args != null && args.Length > 0 && args[0].ToLower() == "runasprogram")
                RunAsProgram();
            else
                RunAsService();
        }

        private static void RunAsProgram()
        {
            try
            {
                new Service1().Run();
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        private static void RunAsService()
        {
            ServiceBase[] ServicesToRun;
            ServicesToRun = new ServiceBase[]
            {
                new Service1()
            };
            ServiceBase.Run(ServicesToRun);
        }
    }
}
