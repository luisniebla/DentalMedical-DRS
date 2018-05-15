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

namespace DentalMedical
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public object Application { get; private set; }
        public static Excel.Application xlApp = new Excel.Application();
        //public static Excel.Workbook xlCBPTemplate = xlApp.Workbooks.Open(@"//serv-az/drs_client_admin/Call Campaigns/ZZ_CBPProccess/CBP Template April (Blackford).xlsx", ReadOnly:true);
        //Excel.Workbook xlCBPData = xlApp.Workbooks.Open(@"\\serv-az\drs_client_admin\Call Campaigns\ZZ_Merge & CBP RawData\CallBackProof Data\042118\Blackford Lifetime CBP.xlsx");

        //Excel.Worksheet xlFoundAppts = xlCBPTemplate.Worksheets[3];
        public MainWindow()
        {
            InitializeComponent();
            xlApp.DisplayAlerts = false;
            
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
            // Create connection string variable. Modify the "Data Source"
            // parameter as appropriate for your environment.
            String sConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + listBoxSearchResults.SelectedItem.ToString() + ";Extended Properties='Excel 12.0;HDR=YES;IMEX=1;';";

            // Create connection object by using the preceding connection string.
            OleDbConnection objConn = new OleDbConnection(sConnectionString);

            // Open connection with the database.
            objConn.Open();

            // The code to follow uses a SQL SELECT command to display the data from the worksheet.

            // Create new OleDbCommand to return data from worksheet.
            OleDbCommand objCmdSelect = new OleDbCommand("SELECT * FROM myRange1", objConn);

            // Create new OleDbDataAdapter that is used to build a DataSet
            // based on the preceding SQL SELECT statement.
            OleDbDataAdapter objAdapter1 = new OleDbDataAdapter();

            // Pass the Select command to the adapter.
            objAdapter1.SelectCommand = objCmdSelect;

            // Create new DataSet to hold information from the worksheet.
            DataSet objDataset1 = new DataSet();

            // Fill the DataSet with the information from the worksheet.
            objAdapter1.Fill(objDataset1, "XLData");

            // Bind data to DataGrid control.
            DataGrid1.ItemsSource = objDataset1.Tables[0].DefaultView;
    

            // Clean up objects.
            objConn.Close();
            //ExcelToMySQL handler = new ExcelToMySQL();
            //String password = TextBoxPassword.Text;
            //String campaign = TextBoxSearchCriteria.Text;

            //handler.ExportExcelHeaders(listBoxSearchResults.SelectedItem.ToString(),@"Apr 2018", @"C:\Users\data\Desktop\" + campaign + "month.csv", password);
            //handler.ExportExcelHeaders(listBoxSearchResults.SelectedItem.ToString(), "Master", @"C:\Users\data\Desktop\" + campaign + "master.csv", password);

            /**
            var dbCon = DBConnection.Instance();
            dbCon.DatabaseName = "test";
            if (dbCon.IsConnect())
            {
                //suppose col0 and col1 are defined as VARCHAR in the DB
                string query = "LOAD DATA LOW_PRIORITY LOCAL INFILE 'C:\\Users\\data\\Desktop\\" + campaign + "month.csv' INTO TABLE `test`.`saltmonth` CHARACTER SET latin1 FIELDS TERMINATED BY '|' LINES TERMINATED BY '\r\n' IGNORE 1 LINES (`First Name`, `Last Name`, `Telephone`, `DIAL`, `Alt #`, `Email`, `DOB/Age`, `Last Visit`, `Resolution`, `Appt Date`, `Appt Time`, `Notes`, `Call Back Date`, `CC Type`, `Provider`, `Insurance`, `Inactives`, `Update`, `# Dials`);";
                var cmd = new MySqlCommand(query, dbCon.Connection);
                var reader = cmd.ExecuteReader();
                //while (reader.Read())
                //{
                //    string someStringFromColumnZero = reader.GetString(0);
                //    string someStringFromColumnOne = reader.GetString(1);
                //    Console.WriteLine(someStringFromColumnZero + "," + someStringFromColumnOne);
                //}
                dbCon.Close();
            }*/
        }
        
        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            Excel.Application xlApp = new Excel.Application();
            xlApp.Visible = true;
            Excel.Workbook xlCBPTemplate = xlApp.Workbooks.Open(@"//serv-az/drs_client_admin/Call Campaigns/ZZ_CBPProccess/CBP Template April (Blackford).xlsx");
            Excel.Worksheet xlFoundAppts = xlCBPTemplate.Worksheets[3];
            Excel.Range lastCBPLine = xlFoundAppts.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);

            // Append found to here.
        }
        
    }
}
