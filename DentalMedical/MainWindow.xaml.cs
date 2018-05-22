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
            ExcelCampaign selectedCampaign = null;
            Excel.Application xlApp = new Excel.Application();
            
            string password = TextBoxPassword.Text;
            string campaign = TextBoxSearchCriteria.Text;
            string filePath = listBoxSearchResults.SelectedItem.ToString();
            
            // Let's try opening this workbook
            try
            {
                selectedCampaign = new ExcelCampaign(xlApp, filePath, password, campaign, "May 2018");
            }
            catch (Exception ex)
            {
                Debug.Write("Could not open excel workbook " + campaign + "\n" + ex.ToString());
                xlApp.Quit();
                return;
            }

            // Our Excel workbook opened correctly. Let's try exporting to csv
            try
            {
                selectedCampaign.ExportHeaders(@"C:\Users\Data\Desktop\");
            }
            catch (Exception ez)
            {
                Debug.Write("Could not write headers " + campaign + "\n" + ez.ToString());
                selectedCampaign.close();
                xlApp.Quit();
                return;
            }
            
            // Okay, we have our csv files. Let's import this baby into SQL

            
        }
       
        private void BtnImport_Click(object sender, RoutedEventArgs e)
        {
            

            // Append found to here.
        }

        private void listBoxSQLTables_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            
        }

        private void App_Close(object sender, System.ComponentModel.CancelEventArgs e)
        {
            
        }
    }
}
