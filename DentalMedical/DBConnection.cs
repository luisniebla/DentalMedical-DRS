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
        private DBConnection()
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

        private static DBConnection _instance = null;
        public static DBConnection Instance()
        {
            if (_instance == null)
                _instance = new DBConnection();
            return _instance;
        }

        public bool IsConnect()
        {
            if (Connection == null)
            {
                if (String.IsNullOrEmpty(databaseName))
                    return false;
                string connstring = string.Format("Server=localhost; database=test; UID=root; password=drs;SslMode=none", databaseName);
                connection = new MySqlConnection(connstring);
                connection.Open();
            }

            return true;
        }

        public MySqlDataReader QueryDB(string table, string command)
        {
            var dbCon = Instance();
            dbCon.DatabaseName = "test";
            if (dbCon.IsConnect())
            {
                var cmd = new MySqlCommand(command, connection);

                var reader = cmd.ExecuteReader();
                //dbCon.Close();
                return reader;
            }

            return null;
        }

        // Close the connection.
        public void Close()
        {
            connection.Close();
        }
    }
}
