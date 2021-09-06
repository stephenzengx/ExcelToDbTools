using Microsoft.EntityFrameworkCore;

namespace ExcelTools
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
