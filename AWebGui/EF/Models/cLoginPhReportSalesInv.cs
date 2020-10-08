using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace AWebGui.EF.Models
{
    [Table("cLoginPhReportSalesInv")]
    public class cLoginPhReportSalesInv
    {
        [Key]
        public int Id { get; set; }
        public string Spv { get; set; }
        public string ExcelFileName { get; set; }
        public int InitDocumentId { get; set; }
        public int cLogId { get; set; }
        public bool Bnd_Activo { get; set; }

        public virtual CLogid cLogin { get; set; }
    }
}
