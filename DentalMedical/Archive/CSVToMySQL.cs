using MySql.Data.MySqlClient;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DentalMedical
{
    class CSVToMySQL
    {
        private MySqlConnection connection;
        private string server;
        private string database;
        private string uid;
        private string password;

        public CSVToMySQL()
        {
            ;
        }

        public void Initialize()
        {
            server = "localhost";
            database = "test";
            uid = "root";
            password = "drs";
            string connectionString;

            // 8.0.8 automatically sets SSLMode on. Need to set to off.
            connectionString = "server=" + server + ";" + "database=" +
            database + ";" + "user id=" + uid + ";" + "password=" + password + ";" + "SslMode=none;" ;

            connection = new MySqlConnection(connectionString);
        }

        public bool OpenConnection()
        {
            try
            {
                connection.Open();
                return true;
            }
            catch (MySqlException ex)
            {
                //When handling errors, you can your application's response based 
                //on the error number.
                //The two most common error numbers when connecting are as follows:
                //0: Cannot connect to server.
                //1045: Invalid user name and/or password.
                switch (ex.Number)
                {
                    case 0:
                        Debug.Write("Cannot connect to server");
                        return false;

                    case 1045:
                        Debug.Write("Bad username/password");
                        return false;
                }
                return false;
            }
        }

        //Close connection
        public bool CloseConnection()
        {
            try
            {
                connection.Close();
                return true;
            }
            catch (MySqlException ex)
            {
                Debug.Write(ex.Message);
                return false;
            }
        }

        public List<string>[] Select(string query, int rowLength, int columnLength)
        {
            //Create a list to store the result
            List<string>[] list = new List<string>[rowLength];

            for (int i = 0; i < rowLength; i++)
                list[i] = new List<string>();

            //Open connection
            if (this.OpenConnection() == true)
            {
                //Create Command
                MySqlCommand cmd = new MySqlCommand(query, connection);
                //Create a data reader and Execute the command
                MySqlDataReader dataReader = cmd.ExecuteReader();

                //Read the data and store them in the list
                //Reads top down
                int row = 0;
                while (dataReader.Read())
                {
                    for (int i = 0; i < columnLength; i++)
                    {
                        if (dataReader.IsDBNull(i))
                            list[row].Add("");
                        else
                            list[row].Add(dataReader.GetString(i));
                    }
                        
                    row++;
                }

                //close Data Reader
                dataReader.Close();

                //close Connection
                this.CloseConnection();

                //return list to be displayed
                return list;
            }
            else
            {
                return list;
            }
        }
    }
}
