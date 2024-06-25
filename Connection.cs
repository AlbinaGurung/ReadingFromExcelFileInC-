using Microsoft.Data.SqlClient;

namespace ReadingExcelData;

public class Connection
{
    string connectionString = "Server=localhost;Database=Student_DB;User Id=SA;Password=rekcod321@; TrustServerCertificate = true";

    public SqlConnection GetDBConnection()
    {
        var connection = new SqlConnection(connectionString);
        connection.Open();
        return connection;
    }
}