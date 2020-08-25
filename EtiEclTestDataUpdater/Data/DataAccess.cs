using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Reflection;
using System.Text;

namespace EtiEclTestDataUpdater.Data
{
    public class DataAccess
    {
        private string _connectionString;
        private string _filePath;
        private string _varianceTemplateFilePath;

        public DataAccess()
        {
            var builder = new ConfigurationBuilder()
                        .SetBasePath(Directory.GetCurrentDirectory())
                        .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true);

            var configuration = builder.Build();

            _connectionString = configuration.GetConnectionString("Default");
            _filePath = configuration.GetConnectionString("TemplateFileFolder");
            _varianceTemplateFilePath = configuration.GetConnectionString("VarianceTemplate");
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
                Console.WriteLine("Error connecting to executing query: " + ex.Message + "\n");
                return false;
            }

        }

        public DataTable GetDataFromQuery(string qry)
        {
            var dt = new DataTable();
            try
            {
                using (SqlConnection con = new SqlConnection(_connectionString))
                {
                    var dataAdapter = new SqlDataAdapter(qry, con);

                    con.Open();
                    dataAdapter.Fill(dt);
                    con.Close();

                    return dt;
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error connecting to executing query: " + ex.Message + "\n");
                Console.ReadKey();
                return dt;
            }
        }

        public string GetFilePath()
        {
            return _filePath;
        }

        public string GetVarianceTemplateFileFolder()
        {
            return _varianceTemplateFilePath;
        }

        public T ParseDataToObject<T>(T t, DataRow row) where T : new()
        {
            // create a new object
            T item = new T();

            // set the item
            SetItemFromRow(item, row);

            // return 
            return item;
        }

        public void SetItemFromRow<T>(T item, DataRow row) where T : new()
        {
            // go through each column
            foreach (DataColumn c in row.Table.Columns)
            {
                // find the property for the column
                PropertyInfo p = item.GetType().GetProperty(c.ColumnName);

                // if exists, set the value
                if (p != null && row[c] != DBNull.Value)
                {
                    p.SetValue(item, row[c], null);
                }
            }
        }
    }
}
