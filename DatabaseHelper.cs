using System.Configuration;
using Microsoft.Data.SqlClient;

namespace AdAgencyManager
{
    public static class DatabaseHelper
    {
        public static string ConnectionString =>
            ConfigurationManager.ConnectionStrings["AdAgencyConnection"]?.ConnectionString;

        public static SqlConnection GetConnection()
        {
            return new SqlConnection(ConnectionString);
        }
    }
}