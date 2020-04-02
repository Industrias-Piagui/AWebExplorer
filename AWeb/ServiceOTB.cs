using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;

namespace AWeb
{
    public partial class ServiceOTB : ServiceBase
    {
        public ServiceOTB()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            Logs.WriteLog("Servicio OTB iniciado");
        }

        protected override void OnStop()
        {
            Logs.WriteLog("Servicio OTB detenido");
        }

        public void Run()
        {

        }

        private void ConfigureTimer()
        {

        }

        //private Task ExecuteAsync()
        //{

        //}
    }
}
