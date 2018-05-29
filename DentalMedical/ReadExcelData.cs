using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DentalMedical
{
    class ReadExcelData
    {

        public ReadExcelData()
        {
            ;
        }

        public void Read(string filePath, string password)
        {
            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                string readerOutput = "";
                ExcelReaderConfiguration xlConf = new ExcelReaderConfiguration();
                xlConf.Password = password;
                xlConf.FallbackEncoding = Encoding.GetEncoding(1252);
                xlConf.AutodetectSeparators = new char[] { ',', ';', '\t', '|', '#' };

                // Auto-detect format, supports:
                //  - Binary Excel files (2.0-2003 format; *.xls)
                //  - OpenXml Excel files (2007 format; *.xlsx)
                var reader = ExcelReaderFactory.CreateReader(stream, xlConf);
                
                // 1. Use the reader methods
                do
                {
                    while (reader.Read())
                    {
                        readerOutput += reader.GetValue(0);
                    }
                    Debug.WriteLine(readerOutput);
                    readerOutput = "";
                } while (reader.NextResult());

                // 2. Use the AsDataSet extension method
                var result = reader.AsDataSet();
                Debug.WriteLine("DONE READING EXCEL");
                // The result of each spreadsheet is in result.Tables
            }
        }
    }
}
