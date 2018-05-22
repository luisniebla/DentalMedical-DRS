using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace DentalMedical
{
    /**
     * Open XML SDK uses the SpreadsheetDocument class to represent an Excel document.
     */
    class OpenXMLSDK
    {
        //SpreadsheetDocument xlDoc;

        public OpenXMLSDK()
        {
            // OpenXML tosses error when attempting to open password protected documents.
            //xlDoc = SpreadsheetDocument.Open(fileName, true);
        }
        // The SAX approach is the recommended way to read a workbook

            /**
        public string DOMRead(string fileName)
        {
            using (SpreadsheetDocument xlDoc = SpreadsheetDocument.Open(fileName, true))
            {
                if (xlDoc == null)
                    return null;
                WorkbookPart xlWkbookPart = xlDoc.WorkbookPart;
                WorksheetPart xlWksheetPart = xlWkbookPart.WorksheetParts.ElementAt(1);
                Worksheet worksheet = xlWksheetPart.Worksheet;



                SheetData sheetData = worksheet.Elements<SheetData>().First();
                string text = "";
                foreach (Row r in sheetData.Elements<Row>())
                {
                    foreach (Cell c in r.Elements<Cell>())
                    {
                        text = c.CellValue.Text;
                        Debug.Write(text + " ");
                    }
                }


                return text;
            }

               
        }

        public string SAXRead()
        {
            if (xlDoc == null)
                return null;

            WorkbookPart workbookPart = xlDoc.WorkbookPart;
            WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();

            OpenXmlReader reader = OpenXmlReader.Create(worksheetPart);
            string text = "";
            while (reader.Read())
            {
                if (reader.ElementType == typeof(Row))
                {
                    reader.ReadFirstChild();
                    do
                    {
                        if (reader.ElementType == typeof(Cell))
                        {
                            Cell c = (Cell)reader.LoadCurrentElement();
                            string cellValue;
                            if (c.DataType != null && c.DataType == CellValues.SharedString)
                            {
                                SharedStringItem ssi = workbookPart.SharedStringTablePart.
                                SharedStringTable.Elements<SharedStringItem>().ElementAt
                                (int.Parse(c.CellValue.InnerText));
                                cellValue = ssi.Text.Text;
                                Debug.WriteLine(cellValue + " ");
                            }
                            else
                            {
                                cellValue = c.CellValue.InnerText;
                                Debug.WriteLine(cellValue + " ");
                            }
                        }

                    } while (reader.ReadNextSibling());
                    Debug.WriteLine("\n");
                }
            }
            
            return text;

        }

        public void LoopRows()
        {
            WorkbookPart workbookPart = xlDoc.WorkbookPart;

            // Iterate through all WorksheetParts
            foreach (WorksheetPart worksheetPart in workbookPart.WorksheetParts)
            {
                OpenXmlPartReader reader = new OpenXmlPartReader(worksheetPart);
                string text;
                string rowNum;
                while (reader.Read())
                {
                    if (reader.ElementType == typeof(Row))
                    {
                        do
                        {
                            if (reader.HasAttributes)
                            {
                                rowNum = reader.Attributes.First(a => a.LocalName == "r").Value;
                                Debug.Write("rowNum: " + rowNum);
                            }

                        } while (reader.ReadNextSibling()); // Skip to the next row

                        break; // We just looped through all the rows so no 
                               // need to continue reading the worksheet
                    }

                    if (reader.ElementType != typeof(Worksheet))
                        reader.Skip();
                }
                reader.Close();
            }

            xlDoc.Close();
        }
    **/
    }
}
