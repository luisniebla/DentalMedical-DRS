using MySql.Data.MySqlClient;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DentalMedical
{
    public class DBConnection
    {
        private string server;
        private string database;
        private string uid;
        private string password;
        public DBConnection()
        {

        }

        private string databaseName = string.Empty;
        public string DatabaseName
        {
            get { return databaseName; }
            set { databaseName = value; }
        }

        public string Password { get; set; }
        private MySqlConnection connection = null;
        public MySqlConnection Connection
        {
            get { return connection; }
        }

        public bool IsConnect()
        {
            if (Connection == null)
            {
                server = "localhost";
                database = "test";
                uid = "root";
                password = "drs";
                string connectionString;

                // 8.0.8 automatically sets SSLMode on. Need to set to off.
                connectionString = "server=" + server + ";" + "database=" +
                database + ";" + "user id=" + uid + ";" + "password=" + password + ";" + "SslMode=none;";

                connection = new MySqlConnection(connectionString);
                connection.Open();
            }

            return true;
        }


        public MySqlDataReader QueryDB(string command)
        {
            try
            {
                if (this.IsConnect())
                {
                    var cmd = new MySqlCommand(command, connection);

                    var reader = cmd.ExecuteReader();
                    //dbCon.Close();

                    return reader;
                }
            }
            catch (MySqlException ex)
            {
                throw ex;
            }
            return null;
            
            
        }

        public MySqlDataReader CreateStringTable(string tableName, string[] columnHeaders)
        {
            if (this.IsConnect())
            {
                string command = "CREATE TABLE " + tableName + " (`" + columnHeaders[0] + "` VARCHAR(255) ";
                for(int i = 1; i < columnHeaders.Length; i++)
                {
                    command += ", `" + columnHeaders[i] + "` VARCHAR(255)";
                }
                command += ");";

                var cmd = new MySqlCommand(command, connection);

                var reader = cmd.ExecuteReader();
                //dbCon.Close();
                return reader;
            }
            return null;
        }
        
        public string ConstructTableSchemaString(string[] columnHeaders, string[] columnDataTypes)
        {
            string command = " (`" + columnHeaders[0] + "` " + columnDataTypes[0];
            for (int i = 1; i < columnHeaders.Length; i++)
            {
                command += ", `" + columnHeaders[i] + "` " + columnDataTypes[i];
            }
            command += ");";

            return command;
        }
        // Close the connection.
        public void Close()
        {
            connection.Close();
        }
    }
}
