using ExcelTools.DbScheme;
using Microsoft.EntityFrameworkCore;

namespace ForExcelImport
{
    public class TestDbContext : DbContext
    {
        // reflect batch resolve fluent-EntityMap Configuration

        /*
          add-migration init
          update-database 
         */
        public TestDbContext(DbContextOptions<TestDbContext> options) : base(options)
        {
        }

        public virtual DbSet<Drug> Drugs { get; set; }

    }
}
