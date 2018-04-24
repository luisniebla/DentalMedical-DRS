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

using Excel = Microsoft.Office.Interop.Excel;

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
        string campaign = "Blackford";
        public MainWindow()
        {
            InitializeComponent();
            xlApp.DisplayAlerts = false;
            
            //ReadExcelHeaders(xlCBPData.Worksheets[4], "CBPData");
            //xlCBPData.Close();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            listBox1.Items.Clear();
            listBox1.UpdateLayout();

            LabelSearchNotification.Content = "Searching...";
            IO myIO = new IO();
            string campaign = TextBoxSearchCriteria.Text;
            string searchCriteria = "*" + campaign + "*Data*.xlsm";
 
            ArrayList foundFiles = myIO.getFiles(@"//serv-az/drs_ofc/05-DENTAL-MEDICAL/", searchCriteria);

            
            
            foreach (string file in foundFiles.ToArray())
            {
                listBox1.Items.Add(file);
            }

            listBox1.UpdateLayout();

            LabelSearchNotification.Content = "DONE";
        }

        private void BtnOpenSelected_Click(object sender, RoutedEventArgs e)
        {
            ExcelManipulation();
        }

        private void ExcelManipulation()
        {
            object refMissing = "";



            // TODO: Programatic finding of CBP using a listview
            // TODO: Programatic finding of Template using a listview
            Excel.Workbook xlDataReport;
            try
            {
                xlDataReport = xlApp.Workbooks.Open(listBox1.SelectedItem.ToString(), Password: "LIFE1", ReadOnly: true);
            }
            catch (System.NullReferenceException)
            {
                xlDataReport = xlApp.Workbooks.Open(listBox1.Items[0].ToString(), Password: "LIFE1", ReadOnly: true);
            }
            xlDataReport.Unprotect();
            ReadExcelHeaders(xlDataReport.Worksheets.get_Item("Master"), "CBPMaster");
            Excel.Worksheet xlMonth;
            try
            {
                xlMonth = xlDataReport.Worksheets.get_Item("April 2018"); // TODO: Month tab can be named differently
            }
            catch (System.Runtime.InteropServices.COMException)
            {
                xlMonth = xlDataReport.Worksheets.get_Item("Apr 2018");
            }
            
            ReadExcelHeaders(xlMonth, "CBPMonth");

            xlDataReport.Close();
            //xlCBPTemplate.Close();

            xlApp.Quit();

            MessageBox.Show("DONE");

            /**
                xlApp.Visible = true;


                Excel.Worksheet xlMaster = xlDataReport.Worksheets[1];
               
                
                
                
                Excel.Worksheet xlData = xlCBPData.Worksheets[2];
                int[] dataColumns = { 1, 2, 11 }; // LName, FName, DOB
                Excel.Range lastCBPLine = xlData.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                Excel.Range lastReportLine = xlMonth.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                
                
                //Excel.Range lastMasterLine = xlMaster.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
                //Excel.Range range = (Excel.Range)xlMaster.Columns["A", Type.Missing];
                
                MessageBox.Show(lastReportLine.Row.ToString());
                // Process the CBP data by cross-refencing with April tab
                // FN    LN  DOB  Res  Note  Last Col
                
                int[] dataReportColumns = { 1, 2, 7, 9, 12, 19 };

                Excel.Range xlCBPDataRange = xlData.UsedRange;

                Excel.Range xlMonthlyData = xlMonth.UsedRange;
        **/

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
