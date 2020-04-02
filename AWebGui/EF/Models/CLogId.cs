using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;

namespace AWebGui.EF.Models
{
    [Table("cLogin")]
    public class CLogid
    {
        [Key]
        [Column("cLogId")]
        public int CLogId { get; set; }
        public string CLogPortal { get; set; }
        public string CLogEmpresa { get; set; }
        public string CLogUsuario { get; set; }
        public string CLogContrasenia { get; set; }
        public string CLogUrl { get; set; }
        public string CLogRutaDescarga { get; set; }
        public string CLogOutUrl { get; set; }
        public string RutaOTB { get; set; }
        public string RutaOBT_SellIn { get; set; }
        public string RutaOTB_SellIn_Plan { get; set; }

    }
}
