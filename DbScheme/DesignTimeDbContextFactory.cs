using System.IO;
using Microsoft.EntityFrameworkCore;
using Microsoft.EntityFrameworkCore.Design;
using Microsoft.Extensions.Configuration;

namespace ForExcelImport
{
    public class DesignTimeDbContextFactory : IDesignTimeDbContextFactory<TestDbContext>
    {
        public TestDbContext CreateDbContext(string[] args)
        {
            var configurationBuilder = new ConfigurationBuilder().SetBasePath(Directory.GetCurrentDirectory()).AddJsonFile("AppSettings.Json");

            var configuration = configurationBuilder.Build();

            var dbContextOptionsBuilder = new DbContextOptionsBuilder<TestDbContext>();

            dbContextOptionsBuilder.UseMySql(configuration.GetConnectionString("ForExcelDb"));



            return new TestDbContext(dbContextOptionsBuilder.Options);
        }
    }
}
