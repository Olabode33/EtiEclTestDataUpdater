using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.IO;
using System.Text;

namespace EtiEclTestDataUpdater.Data
{
    public class DataAccess
    {
        private string _connectionString;
        private string _filePath;

        public DataAccess()
        {
            var builder = new ConfigurationBuilder()
                        .SetBasePath(Directory.GetCurrentDirectory())
                        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            var configuration = builder.Build();

            _connectionString = configuration.GetConnectionString("Default");
            _filePath = configuration.GetConnectionString("TemplateFileFolder");
        }

        public bool ExecuteQuery(string qry)
        {
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    var cmd = new SqlCommand(qry, con);

                    con.Open();
                    var i = cmd.ExecuteNonQuery();
                    con.Close();

                    return true;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error connecting to executing query: " + ex.Message + "\n Query: " + qry);
                return false;
            }

        }

        public string GetFilePath()
        {
            return _filePath;
        }
    }
}
