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

        [HttpPost]
        public IActionResult Post([FromBody] SalesRangeRequestModel range)
        {
            var credentials = (from x in context.CLogid
                               where x.CLogPortal == "PH"
                               select x).FirstOrDefault();
            var filePath = credentials.CLogRutaDescarga;
            if (!filePath.EndsWith(@"\"))
                filePath += @"\";

            RemoveSales(range, filePath);
            var explorer = new AwebExplorer();
            var files = explorer.OnlyDownloadSales(credentials.CLogUsuario, credentials.CLogContrasenia, range.From, range.To);
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

                context.Archivos.Add(new Archivos
                {
                    Fecha = date,
                    Cliente = "3",
                    Archivo = fileName,
                    Ruta = filePath,
                    Tipo = "VTA"
                });
                context.SaveChanges();
            }
        }
    }
}
