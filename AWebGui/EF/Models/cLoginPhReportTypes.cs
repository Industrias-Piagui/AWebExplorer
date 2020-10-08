using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Threading.Tasks;

namespace AWebGui.EF.Models
{
    [Table("cLoginPhReportTypes")]
    public class cLoginPhReportTypes
    {
        [Key]
        public int Id { get; set; }
        public string Nombre { get; set; }
    }
}
