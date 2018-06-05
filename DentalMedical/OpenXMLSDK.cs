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
        
        public string GetCellString(WorkbookPart workbookPart, Worksheet wksht, string columnName, uint rowIndex)
        {
            Cell cell = GetCell(wksht, columnName, rowIndex);
            string cellValue = string.Empty;

            if (cell == null)
                cellValue = "";

            if (cell != null &&  cell.DataType != null)
            {
                if (cell.DataType == CellValues.SharedString)
                {
                    int id = -1;

                    if (Int32.TryParse(cell.InnerText, out id))
                    {
                        SharedStringItem item = GetSharedStringItemById(workbookPart, id);

                        if (item.Text != null)
                        {
                            cellValue = item.Text.Text;
                        }
                        else if (item.InnerText != null)
                        {
                            cellValue = item.InnerText;
                        }
                        else if (item.InnerXml != null)
                        {
                            cellValue = item.InnerXml;
                        }
                    }
                }
            }

            return cellValue;
        }
        public void openLargeFile(string fileName)
        {
            using (SpreadsheetDocument spreadsheetDocument =
                SpreadsheetDocument.Open(fileName, false))
            {
                var sheets = spreadsheetDocument.WorkbookPart.Workbook.Descendants<Sheet>();
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                foreach (Sheet sheet in sheets)
                {
                    WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(sheet.Id);
                    Worksheet worksheet = worksheetPart.Worksheet;

                    SheetData sheetData = worksheetPart.Worksheet.Elements<SheetData>().First();

                    string text;
                    char[] cols = "ABCDEFGHIJKLMNOPQRSTUVWXYZ".ToCharArray();
                    int colIndex = 0;
                    while (colIndex != -1)
                    {
                        foreach (Row r in sheetData.Elements<Row>())
                        {
                            string cellValue = GetCellString(workbookPart, worksheet, cols[colIndex].ToString(), r.RowIndex);

                            Debug.Write(cellValue);
                            
                        }
                        Debug.WriteLine("");
                        colIndex++;
                        string nextHeader = GetCellString(workbookPart, worksheet, cols[colIndex].ToString(), 1);
                        if (nextHeader == "" || colIndex > 20)
                        {
                            colIndex = -1;
                        }
                    }
                    
                    
                }
            }
            
        }

        // Retrieve the value of a cell, given a file name, sheet name, 
        // and address name.
        public static string GetCellValue(string fileName,
            string sheetName,
            string addressName)
        {
            string value = null;

            // Open the spreadsheet document for read-only access.
            using (SpreadsheetDocument document =
                SpreadsheetDocument.Open(fileName, false))
            {
                // Retrieve a reference to the workbook part.
                WorkbookPart wbPart = document.WorkbookPart;

                // Find the sheet with the supplied name, and then use that 
                // Sheet object to retrieve a reference to the first worksheet.
                Sheet theSheet = wbPart.Workbook.Descendants<Sheet>().
                  Where(s => s.Name == sheetName).FirstOrDefault();

                // Throw an exception if there is no sheet.
                if (theSheet == null)
                {
                    throw new ArgumentException("sheetName");
                }

                // Retrieve a reference to the worksheet part.
                WorksheetPart wsPart =
                    (WorksheetPart)(wbPart.GetPartById(theSheet.Id));

                // Use its Worksheet property to get a reference to the cell 
                // whose address matches the address you supplied.
                Cell theCell = wsPart.Worksheet.Descendants<Cell>().
                  Where(c => c.CellReference == addressName).FirstOrDefault();

                // If the cell does not exist, return an empty string.
                if (theCell != null)
                {
                    value = theCell.InnerText;

                    // If the cell represents an integer number, you are done. 
                    // For dates, this code returns the serialized value that 
                    // represents the date. The code handles strings and 
                    // Booleans individually. For shared strings, the code 
                    // looks up the corresponding value in the shared string 
                    // table. For Booleans, the code converts the value into 
                    // the words TRUE or FALSE.
                    if (theCell.DataType != null)
                    {
                        switch (theCell.DataType.Value)
                        {
                            case CellValues.SharedString:

                                // For shared strings, look up the value in the
                                // shared strings table.
                                var stringTable =
                                    wbPart.GetPartsOfType<SharedStringTablePart>()
                                    .FirstOrDefault();

                                // If the shared string table is missing, something 
                                // is wrong. Return the index that is in
                                // the cell. Otherwise, look up the correct text in 
                                // the table.
                                if (stringTable != null)
                                {
                                    value =
                                        stringTable.SharedStringTable
                                        .ElementAt(int.Parse(value)).InnerText;
                                }
                                break;

                            case CellValues.Boolean:
                                switch (value)
                                {
                                    case "0":
                                        value = "FALSE";
                                        break;
                                    default:
                                        value = "TRUE";
                                        break;
                                }
                                break;
                        }
                    }
                }
            }
            return value;
        }
        private static Cell GetCell(Worksheet worksheet, string columnName, uint rowIndex)
        {
            Row row = GetRow(worksheet, rowIndex);

            if (row == null)
                return null;

            if (row.Elements<Cell>() == null)
                return null;

            // Can throw invalid operation exception
            return row.Elements<Cell>().Where(c => string.Compare
                      (c.CellReference.Value, columnName +
                      rowIndex, true) == 0).FirstOrDefault();
        }
        // Given a worksheet and a row index, return the row.
private static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().
                  Elements<Row>().Where(r => r.RowIndex == rowIndex).First();
        }

        public static SharedStringItem GetSharedStringItemById(WorkbookPart workbookPart, int id)
        {
            return workbookPart.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
        }

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
