using AWebGui.EF;
using System.Collections.Generic;
using System.Linq;
using Microsoft.AspNetCore.Mvc;
using AWebGui.Models.Request;
using PortalPhRobot;
using PortalPhRobot.BrowserRobot;
using System.Text.RegularExpressions;
using AWebGui.EF.Models;
using System.IO;
using OfficeOpenXml;
using PortalPhRobot.Extensions;
using PortalPhRobot.Models;
using Newtonsoft.Json;

namespace AWebGui.Controllers
{
    [Route("api/[controller]")]
    public class SalesController : Controller
    {
        private static readonly Regex extractDate = new Regex(@"_\d{8}_");
        private readonly PortalesContext context;

        public SalesController(PortalesContext context)
        {
            this.context = context;
        }

        public class User : DownloadFileModel
        {
            public string DownloadPath { get; set; }
        }

        [HttpPost]
        public IActionResult Post([FromBody] SalesRangeRequestModel range)
        {
            var credentials = (from x in context.CLogid
                               where x.CLogPortal == "PH" && x.CLogId == 10
                               select x).FirstOrDefault();

            List<DownloadFileSalesVsInvParams> LstSInvIDS = new List<DownloadFileSalesVsInvParams>();
            var t1 = (from s in context.cLoginPhReportSalesInvs
                      where s.Bnd_Activo == true && s.cLogId == credentials.CLogId && s.ExcelFileName.Contains("VTA")
                      select new
                      {
                          s.Id,
                          cLogUsuario = credentials.CLogUsuario,
                          s.Spv,
                          s.ExcelFileName,
                          s.InitDocumentId
                      }).ToList();
            foreach (var rf in t1)
            {
                List<ProcesDocumentModel> spv2 = new List<ProcesDocumentModel>();
                spv2 = JsonConvert.DeserializeObject<List<ProcesDocumentModel>>(rf.Spv);
                LstSInvIDS.Add(new DownloadFileSalesVsInvParams
                {
                    cLogUsuario = rf.cLogUsuario,
                    Spv = spv2,
                    ExcelFileName = rf.ExcelFileName,
                    InitDocumentId = rf.InitDocumentId
                });
            }
            User Login = new User();
            Login.IDlogin = credentials.CLogId;
            Login.Password = credentials.CLogContrasenia;
            Login.User = credentials.CLogUsuario;
            Login.DownloadPath = "";
            Login.Variables = (from x1 in context.cLoginPhReportVariables
                               select new DownloadFileVariablesModel
                               {
                                   NavUrl = x1.NavUrl,
                                   ObjIds = x1.ObjIds,
                                   ExcelFileName = x1.ExcelFileName,
                                   TypeId = (DownloadFileVariablesTypes)x1.TypeId
                               }).ToList();
            Login.SalesVsInvParams = LstSInvIDS;
            var filePath =  credentials.CLogRutaDescarga;//@"C:\Users\alejandro_reyes\Desktop\BWPH\";//
            if (!filePath.EndsWith(@"\"))
            {
                filePath += @"\";
            }
            RemoveSales(range, filePath);
            var explorer = new AwebExplorer();
            var files = explorer.OnlyDownloadSales(Login, range.From, range.To);//"";// explorer.OnlyDownloadSales(credentials.CLogUsuario, credentials.CLogContrasenia, range.From, range.To);
            WriteSales(filePath, files);
            return NoContent();
        }

        private void RemoveSales(SalesRangeRequestModel range, string filePath)
        {
            var dates = new List<string>();
            var date = range.From;
            do
            {
                dates.Add(date.ToString("ddMMyyyy"));
                date = date.AddDays(1);
            } while (range.To >= date);

            context.Archivos.RemoveRange(from x in context.Archivos
                                         where dates.Contains(x.Fecha)
                                         && x.Ruta == filePath
                                         && x.Tipo == "VTA"
                                         select x);
            context.SaveChanges();
        }

        private void WriteSales(string filePath, List<SalesVsInventories> files)
        {
            foreach (var file in files)
            {
                var fileName = file.ExcelFileName.Replace(".xlsx", ".csv");
                var date = extractDate.Match(fileName).Groups.FirstOrDefault().Value.Replace("_", "");

                using (var memory = new MemoryStream(file.ExcelContent))
                {
                    using (var excel = new ExcelPackage(memory))
                    {
                        var sheet = excel.Workbook.Worksheets.FirstOrDefault();
                        sheet.DeleteColumn(1);
                        sheet.DeleteRow(1);
                        sheet.DeleteRow(1);
                        var buffer = sheet.GetCsv();
                        System.IO.File.WriteAllBytes($@"{filePath}\{fileName}", buffer);
                    }
                }

                //context.Archivos.Add(new Archivos
                //{
                //    Fecha = date,
                //    Cliente = "3",
                //    Archivo = fileName,
                //    Ruta = filePath,
                //    Tipo = "VTA"
                //});
                //context.SaveChanges();
            }
        }
    }
}
