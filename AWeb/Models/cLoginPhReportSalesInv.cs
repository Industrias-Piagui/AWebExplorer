//------------------------------------------------------------------------------
// <auto-generated>
//     Este código se generó a partir de una plantilla.
//
//     Los cambios manuales en este archivo pueden causar un comportamiento inesperado de la aplicación.
//     Los cambios manuales en este archivo se sobrescribirán si se regenera el código.
// </auto-generated>
//------------------------------------------------------------------------------

namespace AWeb.Models
{
    using System;
    using System.Collections.Generic;
    
    public partial class cLoginPhReportSalesInv
    {
        public int Id { get; set; }
        public string Spv { get; set; }
        public string ExcelFileName { get; set; }
        public int InitDocumentId { get; set; }
        public int cLogId { get; set; }
    
        public virtual cLogin cLogin { get; set; }
    }
}