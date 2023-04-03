using Microsoft.Data.SqlClient;
using Microsoft.EntityFrameworkCore;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;

namespace WorkerCheckDataTKS
{
    public class Data : DbContext
    {
        private readonly IConfiguration _iConfiguration;
        public string _connectionString;
        public Data()
        {
            var uri = new UriBuilder(Assembly.GetExecutingAssembly().CodeBase);
            var path = Uri.UnescapeDataString(uri.Path);
            IConfigurationBuilder builder = new ConfigurationBuilder()
                        .SetBasePath(Path.GetDirectoryName(path))
                        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);
            _iConfiguration = builder.Build();
            _connectionString = _iConfiguration.GetConnectionString("myconn");
        }
        protected override void OnConfiguring(DbContextOptionsBuilder optionsBuilder)
        {
            _ = optionsBuilder.UseSqlServer(_connectionString, providerOptions => providerOptions.CommandTimeout(60000))
                          .UseQueryTrackingBehavior(QueryTrackingBehavior.NoTracking);
        }
       
    }
}
