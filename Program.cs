using System.Data;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ReadingExcelData;
using Microsoft.Data.SqlClient;
using static ReadingExcelData.DTOExcelFile;
using static ReadingExcelData.Connection;

namespace ReadingExcelData
{
    class Program
    {
        static void Main(string[] args)
        {
            string filePath = @"/Users/albinagurung/StdData.xlsx";
            List<StudentRecordDTO> excelData = new List<StudentRecordDTO>();
            if (File.Exists(filePath))
            {
                try
                {
                   excelData = ReadExcelFile(filePath);
                    Console.WriteLine("Excel Data:");
                    foreach (var record in excelData)
                    {
                        Console.WriteLine(
                            $"{record.RegNo}\t{record.FirstName}\t{record.LastName}\t{record.RollNo}\t{record.Class}\t{record.Group}\t{record.Section}\t{record.DateOfBirth}\t{record.Sex}\t{record.FFirstName}\t{record.FLastName}\t{record.MFirstName}\t{record.MLastName}\t{record.MobileNo}\t{record.PhoneR}\t{record.MMobileNo}\t{record.Zone}\t{record.Province}\t{record.District}\t{record.Municipality}\t{record.WardNo}\t{record.Tole}\t{record.Dues}\t{record.Balance}");
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error occurred while reading the Excel file");
                    Console.WriteLine(ex.Message);
                }
            }
            else
            {
                Console.WriteLine($"The file at path {filePath} does not exist");
            }

            //Inserting bulk data into Students using SQLbulkCopy
            Connection con = new Connection();

            using (SqlConnection dbConnection = con.GetDBConnection())
            {
                Console.WriteLine("Connection successfully established");
                using (SqlBulkCopy bulkCopy = new SqlBulkCopy(dbConnection))
                {
                    bulkCopy.DestinationTableName = "dbo.Students";
                    DataTable dt = ConvertToDataTable(excelData);
                    try
                    {
                        bulkCopy.WriteToServer(dt);
                        Console.WriteLine("Data successfully inserted into the database.");
                    
                    }
                    catch (Exception e)
                    {
                        Console.WriteLine(e);
                        Console.WriteLine("Error occurred while inserting data into the database.");
                        
                        throw;
                    }
                }
            }
            Console.WriteLine("Press any key to exit.");
            Console.ReadKey();

        }

        private static DataTable ConvertToDataTable(List<DTOExcelFile> excelData)
        {
            DataTable table = new DataTable();

            // Define the columns
            table.Columns.Add("RegNo", typeof(int));
            table.Columns.Add("FirstName", typeof(string));
            table.Columns.Add("LastName", typeof(string));
            table.Columns.Add("RollNo", typeof(int));
            table.Columns.Add("Class", typeof(string));
            table.Columns.Add("Group", typeof(string));
            table.Columns.Add("Section", typeof(string));
            table.Columns.Add("DateOfBirth", typeof(string)); // Assuming DateOfBirth is a string in your DTO
            table.Columns.Add("Sex", typeof(string));
            table.Columns.Add("FFirstName", typeof(string));
            table.Columns.Add("FLastName", typeof(string));
            table.Columns.Add("MFirstName", typeof(string));
            table.Columns.Add("MLastName", typeof(string));
            table.Columns.Add("MobileNo", typeof(string));
            table.Columns.Add("PhoneR", typeof(string));
            table.Columns.Add("MMobileNo", typeof(string));
            table.Columns.Add("Zone", typeof(string));
            table.Columns.Add("Province", typeof(string));
            table.Columns.Add("District", typeof(string));
            table.Columns.Add("Municipality", typeof(string));
            table.Columns.Add("WardNo", typeof(int));
            table.Columns.Add("Tole", typeof(string));
            table.Columns.Add("Dues", typeof(decimal));
            table.Columns.Add("Balance", typeof(decimal));
            foreach (var item in excelData)
            {
                DataRow row = table.NewRow();
                row["RegNo"] = item.RegNo;
                row["FirstName"] = item.FirstName;
                row["LastName"] = item.LastName;
                row["RollNo"] = item.RollNo;
                row["Class"] = item.Class;
                row["Group"] = item.Group;
                row["Section"] = item.Section;
                row["DateOfBirth"] = item.DateOfBirth;
                row["Sex"] = item.Sex;
                row["FFirstName"] = item.FFirstName;
                row["FLastName"] = item.FLastName;
                row["MFirstName"] = item.MFirstName;
                row["MLastName"] = item.MLastName;
                row["MobileNo"] = item.MobileNo;
                row["PhoneR"] = item.PhoneR;
                row["MMobileNo"] = item.MMobileNo;
                row["Zone"] = item.Zone;
                row["Province"] = item.Province;
                row["District"] = item.District;
                row["Municipality"] = item.Municipality;
                row["WardNo"] = item.WardNo;
                row["Tole"] = item.Tole;
                row["Dues"] = item.Dues;
                row["Balance"] = item.Balance;

                table.Rows.Add(row);
            }
            return table;
            
        }

        static List<StudentRecordDTO> ReadExcelFile(string filePath)
        {
            List<StudentRecordDTO> data = new List<StudentRecordDTO>();
            try
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
                {
                    var workbookPart = document.WorkbookPart;
                    var worksheetPart = workbookPart.WorksheetParts.FirstOrDefault();
                    if (worksheetPart != null)
                    {
                        var sheetData = worksheetPart.Worksheet.Elements<SheetData>().FirstOrDefault();
                        if (sheetData != null)
                        {
                            foreach (var row in sheetData.Elements<Row>().Skip(1))
                            {
                                var cells = row.Elements<Cell>().ToArray();
                                if (cells.Length < 24)
                                {
                                    Console.WriteLine(
                                        $"Row {row.RowIndex} has insufficient cells. Expected 24, found {cells.Length}");
                                    continue;
                                }

                                var cellDictionary = cells.ToDictionary(cell => (int)GetColumnIndex(cell.CellReference),
                                    cell => cell);

                                string regNoValue = GetCellValue(document,
                                    cellDictionary.ContainsKey(0) ? cellDictionary[0] : null);
                                if (!int.TryParse(regNoValue, out int regNo))
                                {
                                    Console.WriteLine(
                                        $"Invalid integer value '{regNoValue}' at row {row.RowIndex}, column 1");
                                    continue;
                                }

                                string rollNoValue = GetCellValue(document,
                                    cellDictionary.ContainsKey(3) ? cellDictionary[3] : null);
                                if (!int.TryParse(rollNoValue, out int rollNo))
                                {
                                    Console.WriteLine(
                                        $"Skipping row {row.RowIndex} as RollNo '{rollNoValue}' is not a valid integer.");
                                    continue;
                                }

                                try
                                {
                                    var record = new StudentRecordDTO
                                    {
                                        RegNo = regNo,
                                        FirstName = GetCellValue(document,
                                            cellDictionary.ContainsKey(1) ? cellDictionary[1] : null),
                                        LastName = GetCellValue(document,
                                            cellDictionary.ContainsKey(2) ? cellDictionary[2] : null),
                                        RollNo = rollNo,
                                        Class = GetCellValue(document,
                                            cellDictionary.ContainsKey(4) ? cellDictionary[4] : null),
                                        Group = GetCellValue(document,
                                            cellDictionary.ContainsKey(5) ? cellDictionary[5] : null),
                                        Section = GetCellValue(document,
                                            cellDictionary.ContainsKey(6) ? cellDictionary[6] : null),
                                        DateOfBirth = GetCellValue(document,
                                            cellDictionary.ContainsKey(7) ? cellDictionary[7] : null),
                                        Sex = GetCellValue(document,
                                            cellDictionary.ContainsKey(8) ? cellDictionary[8] : null),
                                        FFirstName = GetCellValue(document,
                                            cellDictionary.ContainsKey(9) ? cellDictionary[9] : null),
                                        FLastName = GetCellValue(document,
                                            cellDictionary.ContainsKey(10) ? cellDictionary[10] : null),
                                        MFirstName = GetCellValue(document,
                                            cellDictionary.ContainsKey(11) ? cellDictionary[11] : null),
                                        MLastName = GetCellValue(document,
                                            cellDictionary.ContainsKey(12) ? cellDictionary[12] : null),
                                        MobileNo = GetCellValue(document,
                                            cellDictionary.ContainsKey(13) ? cellDictionary[13] : null),
                                        PhoneR = GetCellValue(document,
                                            cellDictionary.ContainsKey(14) ? cellDictionary[14] : null),
                                        MMobileNo = GetCellValue(document,
                                            cellDictionary.ContainsKey(15) ? cellDictionary[15] : null),
                                        Zone = GetCellValue(document,
                                            cellDictionary.ContainsKey(16) ? cellDictionary[16] : null),
                                        District = GetCellValue(document,
                                            cellDictionary.ContainsKey(17) ? cellDictionary[17] : null),
                                        Province = GetCellValue(document,
                                            cellDictionary.ContainsKey(18) ? cellDictionary[18] : null),
                                        Municipality = GetCellValue(document,
                                            cellDictionary.ContainsKey(19) ? cellDictionary[19] : null),
                                        WardNo = ParseInt(
                                            GetCellValue(document,
                                                cellDictionary.ContainsKey(20) ? cellDictionary[20] : null),
                                            row.RowIndex, 20),
                                        Tole = GetCellValue(document,
                                            cellDictionary.ContainsKey(21) ? cellDictionary[21] : null),
                                        Dues = ParseDecimal(
                                            GetCellValue(document,
                                                cellDictionary.ContainsKey(22) ? cellDictionary[22] : null),
                                            row.RowIndex, 22),
                                        Balance = ParseDecimal(
                                            GetCellValue(document,
                                                cellDictionary.ContainsKey(23) ? cellDictionary[23] : null),
                                            row.RowIndex, 23)
                                    };
                                    data.Add(record);
                                }
                                catch (Exception ex)
                                {
                                    Console.WriteLine($"Error processing row {row.RowIndex}: {ex.Message}");
                                }
                            }
                        }
                        else
                        {
                            Console.WriteLine("SheetData element is null");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("An error occurred while reading the Excel file.");
                Console.WriteLine(ex.Message);
            }

            return data;
        }

        static uint GetColumnIndex(string cellReference)
        {
            string columnReference = new string(cellReference.Where(Char.IsLetter).ToArray());
            uint columnIndex = 0;
            foreach (char c in columnReference)
            {
                columnIndex = (uint)(columnIndex * 26 + (c - 'A' + 1));
            }

            return columnIndex - 1;
        }

        static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            if (cell == null || cell.CellValue == null)
            {
                return "";
            }

            string value = cell.CellValue.Text;

            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
            {
                SharedStringTablePart sstpart =
                    document.WorkbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                return sstpart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;
            }

            return value;
        }

        // validation 
        static int ParseInt(string value, uint rowIndex, int columnIndex)
        {
            if (int.TryParse(value, out int result))
            {
                return result;
            }

            throw new FormatException($"Invalid integer value '{value}' at row {rowIndex}, column {columnIndex + 1}");
        }

        static decimal ParseDecimal(string value, uint rowIndex, int columnIndex)
        {
            if (decimal.TryParse(value, out decimal result))
            {
                return result;
            }

            throw new FormatException($"Invalid decimal value '{value}' at row {rowIndex}, column {columnIndex + 1}");
        }
    }
}