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
            ExcelToMySQL handler = new ExcelToMySQL();
            String password = TextBoxPassword.Text;
            String campaign = TextBoxSearchCriteria.Text;

            //ArrayList monthHeaders = handler.ExportExcelHeaders(listBoxSearchResults.SelectedItem.ToString(),3, @"C:\Users\data\Desktop\" + campaign + "month.csv", password);
            //ArrayList masterHeaders = handler.ExportExcelHeaders(listBoxSearchResults.SelectedItem.ToString(), 1, @"C:\Users\data\Desktop\" + campaign + "master.csv", password);


            var dbCon = DBConnection.Instance();
            dbCon.DatabaseName = "test";
            if (dbCon.IsConnect())
            {
                var query = "SELECT * FROM logmeins WHERE id = 6;";

                var cmd = new MySqlCommand(query, dbCon.Connection);

                var reader = cmd.ExecuteReader();

                while (reader.Read())
                {
                    string campaignInfo = "";
                    for (int i = 1; i < reader.FieldCount; i++)
                        campaignInfo += listBoxSQLTables.Items.Add(reader.GetString(i));
                    TextBlockCampaignInfo.Text = campaignInfo;
                }
                dbCon.Close();
                listBoxSearchResults.UpdateLayout();
                MessageBox.Show("SUCCESS");

            }
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

        private void listBoxSQLTables_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
        }
    }
}
