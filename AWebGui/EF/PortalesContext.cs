using AWebGui.EF.Models;
using Microsoft.EntityFrameworkCore;

namespace AWebGui.EF
{
    public class PortalesContext : DbContext
    {
        public DbSet<Archivos> Archivos { get; set; }
        public DbSet<CLogid> CLogid { get; set; }

        public PortalesContext(DbContextOptions<PortalesContext> options) : base(options)
        { }

        protected override void OnModelCreating(ModelBuilder modelBuilder)
        {
            modelBuilder.Entity<Archivos>().HasKey(k => new { k.Fecha, k.Cliente, k.Archivo });
        }
    }
}
