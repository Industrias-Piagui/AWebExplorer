using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace AWebGui.EF.Models
{
    [Table("cLoginPhReportVariables")]
    public class cLoginPhReportVariables
    {
        [Key]
        public int Id { get; set; }
        public string NavUrl { get; set; }
        public string ObjIds { get; set; }
        public string ExcelFileName { get; set; }
        public int TypeId { get; set; }

        //public virtual cLoginPhReportTypes cLoginPhReportTypes { get; set; }
    }
}
