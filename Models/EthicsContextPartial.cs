using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Configuration;

namespace ExcelReport.Models
{
    public partial class EthicsContext
    {
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            // Only configure if the options haven't been set yet.
            if (!optionsBuilder.IsConfigured)
            {
                // Grab the connection string from your app's static Configuration property
                string connectionString = ExcelReport.App.Configuration.GetConnectionString("SqliteConnection");

                optionsBuilder.UseSqlite(connectionString);
            }
        }
    }
}
