using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Composition;
using System.Linq;
using System.Threading.Tasks;

namespace ExcelExport.Models
{
    public class ExportContext : DbContext
    {
        public ExportContext(DbContextOptions<ExportContext> contextOptions) : base(contextOptions)
        {
        }
        public DbSet<Employees> Employees { get; set; }
    }
}
