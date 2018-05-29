using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

using System.IO;
using System.Text.RegularExpressions;
using System.Collections;

using System.Data.OleDb;
using System.Data;

using Excel = Microsoft.Office.Interop.Excel;
using ExcelDataReader;
using System.Data;
using MySql.Data.MySqlClient;
using System.Runtime.InteropServices;
using System.Data.Common;
using System.Diagnostics;
using Microsoft.Office.Interop.Excel;



namespace DentalMedical
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : System.Windows.Window
    {
        public object Application { get; private set; }
        //public static Excel.Application xlApp = new Excel.Application();
        //public static Excel.Workbook xlCBPTemplate = xlApp.Workbooks.Open(@"//serv-az/drs_client_admin/Call Campaigns/ZZ_CBPProccess/CBP Template April (Blackford).xlsx", ReadOnly:true);
        //Excel.Workbook xlCBPData = xlApp.Workbooks.Open(@"\\serv-az\drs_client_admin\Call Campaigns\ZZ_Merge & CBP RawData\CallBackProof Data\042118\Blackford Lifetime CBP.xlsx");

        //Excel.Worksheet xlFoundAppts = xlCBPTemplate.Worksheets[3];
        public MainWindow()
        {
            InitializeComponent();
            //xlApp.DisplayAlerts = false;

            //ReadExcelHeaders(xlCBPData.Worksheets[4], "CBPData");
            //xlCBPData.Close();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            listBoxSearchResults.Items.Clear();
            listBoxSearchResults.UpdateLayout();

            IO myIO = new IO();

            string campaign = TextBoxSearchCriteria.Text;
            string searchCriteria = "*" + campaign + "*Data*.xlsm";

            ArrayList foundFiles = myIO.getFiles(@"//serv-az/drs_ofc/05-DENTAL-MEDICAL/", searchCriteria);

            foreach (string file in foundFiles.ToArray())
            {
                listBoxSearchResults.Items.Add(file);
            }

            listBoxSearchResults.UpdateLayout();

            LabelSearchNotification.Content = "DONE";
        }


        private void BtnOpenSelected_Click(object sender, RoutedEventArgs e)
        {
            //Excel.Application xlApp = new Excel.Application();

            string password = TextBoxPassword.Text;
            string campaign = TextBoxSearchCriteria.Text;
            string filePath = listBoxSearchResults.SelectedItem.ToString();

            // Excel.Application is too slow. Let's try something else.
            Excel.Application xlApp = new Excel.Application();

            PIMACampaign thc = new PIMACampaign(xlApp, password, filePath, campaign, "May-Merge 2018");

            Debug.WriteLine(thc.HeadersToString());

            thc.AttemptCallBackProof();
            thc.close();

            xlApp.Quit();

        }

        // Okay, we have our csv files. Let's import this baby into SQL
        //

        //string[] cols = { "First Name","Last Name","Telephone","DIAL","Alt#","Email","DOB/Age","Last Visit","Resolution","Date","Time","Notes","Call Back Date","CC Type","Provider","Insurance","Update","CBP Date" };
        //MySqlDataReader reader2 = dbhandler.CreateStringTable("Diablo_Month", cols);
        // reader2.Close();

        /**
        MySqlDataReader reader = dbhandler.QueryDB(@"LOAD DATA INFILE 'C:\\\Users\\\Data\\\Desktop\\\DiabloMonth.csv' INTO TABLE test.Diablo_Month FIELDS TERMINATED BY ',' IGNORE 2 LINES;");
        while (reader.Read())
        {
            for(int i = 0; i < reader.FieldCount; i++)
            {
                Debug.WriteLine(reader.GetString(i));
            }
        }
        **/


        //List<string>[] output = dbconnection.Select("SELECT * FROM logmeins", 39, 26 );

        //List<string>[] output = dbconnection.Select(@"LOAD DATA INFILE 'C:\Users\data\Desktop\DiabloMaster.csv' INTO TABLE test.")
        //for (int i = 0; i < output.Length; i++)
        //{
        //    MessageBox.Show(output[i].ToArray()[0]);
        //}
        //xlApp.Quit();
   
        public void oledb()
        {
            using (OleDbConnection conn = new OleDbConnection())
            {
                System.Data.DataTable dt = new System.Data.DataTable();
                string Import_FileName = @"C:\Users\data\Desktop\exceloutput.xlsx";
                string fileExtension = System.IO.Path.GetExtension(Import_FileName);
                if (fileExtension == ".xls")
                    conn.ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 8.0;HDR=YES;'";
                if (fileExtension == ".xlsx")
                    conn.ConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + Import_FileName + ";" + "Extended Properties='Excel 12.0 Xml;HDR=YES;'";


                using (OleDbCommand comm = new OleDbCommand())
                {
                    comm.CommandText = "Select * from [" + "Master" + "$]";

                    comm.Connection = conn;

                    using (OleDbDataAdapter da = new OleDbDataAdapter())
                    {
                        da.SelectCommand = comm;
                        da.Fill(dt);
                    }

                }
            }
        }
        private void BtnImport_Click(object sender, RoutedEventArgs e)
        { 
            // Append found to here.
        }

        private void listBoxSQLTables_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            // This isn't being used currently but I do think it will be used in the future.
        }

        private void App_Close(object sender, System.ComponentModel.CancelEventArgs e)
        {
            
        }
    }
}
