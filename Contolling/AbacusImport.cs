using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Controlling
{
    public class Booking
    {
        public string Employee{ get; init; }
        public DateOnly Date { get; init; }
        public string WorkType { get; init; }
        public string WorkDescription { get; init; }
        public float Hours { get; init; }
        public float Rate { get; init; }
        public string TicketId { get; set; }
        public Contract Contract { get; init; }
        public override string ToString()
        {
            return $"Project: {Contract.Project.Name}, Employee: {Employee}, Date: {Date}, " +
                   $"Type: {WorkType}, Desc: {WorkDescription}, " +
                   $"Hours: {Hours}, Rate: {Rate}, " +
                   $"TicketId: {TicketId}";
        }
    }
    internal class AbacusImport
    {

        static string[] columns = { "Projekt-Nr.", "Mitarbeiter", "Datum", "Leistung", "Beschreibung", "V", "Abrechnung", "Anzahl", "Ansatz", "Betrag", "CR-Nr." };
        public static IEnumerable<Booking> ParseExcelFile(string fileName, ProjectSettings settings)
        {
            var abacusData = new List<Booking>();
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(fileName, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().FirstOrDefault();
                WorksheetPart worksheetPart = (WorksheetPart)(workbookPart.GetPartById(sheet.Id));
                Worksheet worksheet = worksheetPart.Worksheet;

                foreach (Row row in worksheet.Descendants<Row>().Skip(2))
                {
                    var cells = row.Descendants<Cell>().Take(columns.Length).ToArray();
                    if (cells.Length == 0) continue;
                    //var firstCell = GetCellValue(cells[ColumnAttribute.GetColumn<AbacusData>(x => x.ProjectNr)], workbookPart);
                    var firstCell = GetCellValue(cells[0], workbookPart);
                    if (string.IsNullOrEmpty(firstCell)) continue;
                    if (firstCell == "Total Stunden Mitarbeiter STD") break;
                    var data = new Booking
                    {
                        Contract = FindProject(GetCellValue(cells[0], workbookPart), settings.Contracts),
                        Employee = GetCellValue(cells[1], workbookPart),
                        //Date = DateTime.ParseExact(GetCellValue(cells[2], workbookPart), "dd.MM.yyyy", CultureInfo.InvariantCulture),
                        Date = DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(cells[2], workbookPart)))),
                        WorkType = GetCellValue(cells[3], workbookPart),
                        WorkDescription = GetCellValue(cells[4], workbookPart),
                        Hours = float.Parse(GetCellValue(cells[7], workbookPart), CultureInfo.InvariantCulture),
                        Rate = float.Parse(GetCellValue(cells[8], workbookPart), CultureInfo.InvariantCulture),
                        TicketId = GetCellValue(cells[10], workbookPart)
                    };
                    abacusData.Add(data);
                    Console.WriteLine();
                    //for (int i = 0; i < cells.Length; i++)
                    //{
                    //    //string text = cells[i].InnerText;
                    //    string text = GetCellValue(cells[i], workbookPart);
                    //}
                }
            }
            return abacusData;
        }

        private static Contract FindProject(string projectNr, IEnumerable<Contract> contracts)
        {
            return contracts.FirstOrDefault(c => c.Id == projectNr);
        }

        //public static DataTable ParseExcelFile2(string fileName)
        //{
        //    DataTable table = new DataTable();

        //    // Define the columns in the DataTable based on the MyData class
        //    table.Columns.Add("ProjectNr", typeof(string));
        //    table.Columns.Add("Employee", typeof(string));
        //    table.Columns.Add("Date", typeof(DateTime));
        //    table.Columns.Add("Leistung", typeof(string));
        //    table.Columns.Add("V", typeof(string));
        //    table.Columns.Add("Abrechnung", typeof(string));
        //    table.Columns.Add("Anzahl", typeof(float));
        //    table.Columns.Add("Ansatz", typeof(float));
        //    table.Columns.Add("Betrag", typeof(float));
        //    table.Columns.Add("TicketId", typeof(string));


        //    string connectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fileName + ";Extended Properties=\"Excel 12.0;HDR=YES;IMEX=1\";";
        //    using (OleDbConnection connection = new OleDbConnection(connectionString))
        //    {
        //        connection.Open();
        //        string sheetName = "Leistungsauszug";
        //        OleDbDataAdapter dataAdapter = new OleDbDataAdapter("SELECT * FROM [" + sheetName + "]", connection);
        //        DataSet dataSet = new DataSet();
        //        dataAdapter.Fill(dataSet);
        //        // Use the DataTable's Merge method to combine the data from the Excel file with the strongly-typed DataTable
        //        table.Merge(dataSet.Tables[0]);
        //    }

        //    return table;
        //}

        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            string value = cell.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                SharedStringTablePart sharedStringPart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                if (sharedStringPart != null)
                {
                    value = sharedStringPart.SharedStringTable.ElementAt(int.Parse(value)).InnerText;
                }
            }
            return value;
        }
    }
}