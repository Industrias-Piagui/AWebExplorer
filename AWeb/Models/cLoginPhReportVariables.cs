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
    
    public partial class cLoginPhReportVariables
    {
        public int Id { get; set; }
        public string NavUrl { get; set; }
        public string ObjIds { get; set; }
        public string ExcelFileName { get; set; }
        public int TypeId { get; set; }
    
        public virtual cLoginPhReportTypes cLoginPhReportTypes { get; set; }
    }
}