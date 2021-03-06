﻿using AWeb.Extensions;
using AWeb.Models;
using OfficeOpenXml;
using PortalPhRobot;
using System;
using System.Data;
using System.Data.Entity;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Timers;
using System.Threading.Tasks;
using PortalPhRobot.BrowserRobot;
using System.Collections.Generic;
using PortalPhRobot.Exceptions;
using System.Threading;
using System.Configuration;
using PortalPhRobot.Models;
using Newtonsoft.Json;

namespace AWeb
{
    public partial class Service1 : ServiceBase
    {
        private TimeSpan timeToRun;
        private System.Timers.Timer timer;

        public Service1()
        {
            InitializeComponent();
        }

        public void Run()
        {
            DownloadFilesAsync().Wait();
        }

        protected override void OnStart(string[] args)
        {
            timeToRun = TimeSpan.Parse(ConfigurationManager.AppSettings["TimeToRun"]);
            ConfigureService();
            Logs.WriteLog("Servicio iniciado");
        }

        protected override void OnStop()
        {
            Logs.WriteLog("Servicio detenido");
        }

        private void ConfigureService()
        {
            var now = DateTime.Now;
            var timeToFire = new DateTime(now.Year, now.Month, now.Day, timeToRun.Hours, timeToRun.Minutes, timeToRun.Seconds);
            if (timeToFire < now)
                timeToFire = timeToFire.AddDays(1);

            var timediff = timeToFire - now;
            timer = new System.Timers.Timer(timediff.TotalMilliseconds)
            {
                AutoReset = false,
                Enabled = true
            };
            timer.Elapsed += Timer_Elapsed;
            timer.Start();
            Logs.WriteLog($"Evento configurado para lanzarse en {timediff.TotalSeconds} segundos");
        }

        private void Timer_Elapsed(object sender, ElapsedEventArgs e)
        {
            Logs.WriteLog("Evento disparado");
            while (true)
            {
                try
                {
                    DownloadFilesAsync().Wait();
                    timer.Stop();
                    timer.Enabled = false;
                    timer.Elapsed -= Timer_Elapsed;
                    timer.Dispose();
                    timer = null;
                    ConfigureService();
                    break;
                }
                catch (WrongLoginException ex)
                {
                    Logs.WriteErrorLog(ex.Message);
                    Thread.Sleep(600_000);
                }
                catch (ServerErrorException ex)
                {
                    Logs.WriteErrorLog(ex.Message);
                    Thread.Sleep(600_000);
                }
                catch (Exception ex)
                {
                    Logs.WriteErrorLog(ex.ToString());
                    Thread.Sleep(600_000);
                }
            }
        }

        private async Task DownloadFilesAsync()
        {
            using (var context = new PortalesPRODEntities())
            {
                var aweb = new AwebExplorer();
                var yesterday = DateTime.Today.AddDays(-1);
                var yesterdayStr = yesterday.ToString("ddMMyyyy");
                var users = await (from x in context.cLogin
                                   where x.cLogPortal == "PH"
                                   select new User
                                   {
                                       User = x.cLogUsuario,
                                       Password = x.cLogContrasenia,
                                       DownloadPath = x.cLogRutaDescarga,
                                       Variables = (from x1 in context.cLoginPhReportVariables
                                                    select new DownloadFileVariablesModel
                                                    {
                                                        NavUrl = x1.NavUrl,
                                                        ObjIds = x1.ObjIds,
                                                        ExcelFileName = x1.ExcelFileName,
                                                        TypeId = (DownloadFileVariablesTypes)x1.TypeId
                                                    }).ToList(),
                                       SalesVsInvParams = (from s in context.cLoginPhReportSalesInv
                                                           join u in context.cLogin on s.cLogId equals u.cLogId
                                                           where u.cLogUsuario == x.cLogUsuario
                                                           select new DownloadFileSalesVsInvParams
                                                           {
                                                               Spv = s.Spv,
                                                               ExcelFileName = s.ExcelFileName,
                                                               InitDocumentId = s.InitDocumentId
                                                           })
                                                           .ToList()
                                   }).ToListAsync();

                foreach (var login in users)
                {
#if !(DEBUG)
                    var downloadPath = login.DownloadPath;
                    var salesList = await aweb.DownloadAsync(login, yesterday, yesterday);
#else
                    var downloadPath = "EXCEL";
                    var salesList = aweb.Download(login, yesterday, yesterday);
#endif
                    MergeInvIpiFiles(ref salesList);

                    foreach (var sales in salesList)
                    {
                        var fileName = sales.ExcelFileName.Replace(".xlsx", ".csv");
                        using (var memory = new MemoryStream(sales.ExcelContent))
                        {
                            using (var excel = new ExcelPackage(memory))
                            {
                                RemoveUnecesaryFields(excel);
                                var buffer = excel.GetCsv(excel.Workbook.Worksheets.FirstOrDefault());
                                File.WriteAllBytes($@"{downloadPath}\{fileName}", buffer);
                            }
                        }

#if !(DEBUG)
                        if (context.ARCHIVOS.FirstOrDefault(x => x.FECHA == yesterdayStr && x.CLIENTE == "3" && x.ARCHIVO == fileName) == null)
                        {
                            context.ARCHIVOS.Add(new ARCHIVOS
                            {
                                FECHA = yesterdayStr,
                                CLIENTE = "3",
                                ARCHIVO = fileName,
                                RUTA = downloadPath.EndsWith(@"\") ? downloadPath : $@"{downloadPath}\",
                                TIPO = fileName.Contains("_VTA_") ? "VTA" : "INV"
                            });
                        }
#endif
                    }

                    context.SaveChanges();
                }
            }
        }

        private void RemoveUnecesaryFields(ExcelPackage excel)
        {
            var sheet = excel.Workbook.Worksheets.FirstOrDefault();
            sheet.DeleteColumn(1);
            sheet.DeleteRow(1);
            sheet.DeleteRow(1);
        }

        private void MergeInvIpiFiles(ref List<SalesVsInventories> salesList)
        {
            var invIpi = (from x in salesList where x.ExcelFileName.Contains("_IPI") select x).ToList();
            SalesVsInventories invIpi1;
            SalesVsInventories invIpi2;
            if (invIpi.FirstOrDefault().ExcelFileName.Contains("_IPI1"))
            {
                invIpi1 = invIpi[0];
                invIpi2 = invIpi[1];
            }
            else
            {
                invIpi1 = invIpi[1];
                invIpi2 = invIpi[0];
            }

            using (var memory1 = new MemoryStream(invIpi1.ExcelContent))
            {
                using (var memory2 = new MemoryStream(invIpi2.ExcelContent))
                {
                    using (var excel1 = new ExcelPackage(memory1))
                    {
                        using (var excel2 = new ExcelPackage(memory2))
                        {
                            using (var excel3 = new ExcelPackage())
                            {
                                var sheet1 = excel1.Workbook.Worksheets.FirstOrDefault();
                                var sheet2 = excel2.Workbook.Worksheets.FirstOrDefault();
                                var sheet3 = excel3.Workbook.Worksheets.Add(sheet1.Name);
                                sheet2.DeleteRow(1);
                                sheet2.DeleteRow(1);
                                sheet3.MergeSheets(sheet1);
                                sheet3.MergeSheets(sheet2);
                                invIpi1.ExcelContent = excel3.GetAsByteArray();
                                salesList.Remove(invIpi2);
                            }

                            /*var sheet = excel2.Workbook.Worksheets.FirstOrDefault();
                            sheet.DeleteRow(1);
                            sheet.DeleteRow(1);
                            excel1.Workbook.Worksheets.FirstOrDefault().MergeSheets(sheet);
                            invIpi1.ExcelContent = excel1.GetAsByteArray();
                            salesList.Remove(invIpi2);*/
                        }
                    }
                }
            }
        }
    }
}
