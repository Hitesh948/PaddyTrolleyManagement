using Microsoft.Data.SqlClient;

namespace PaddyTrolleyManagement
{
    public static class DatabaseHelper
    {
        private static readonly string connectionString =
            @"Server=(localdb)\MSSQLLocalDB;Database=PaddyTrolleyDB;Trusted_Connection=True;TrustServerCertificate=True;";

        public static SqlConnection GetConnection()
        {
            return new SqlConnection(connectionString);
        }
    }
}