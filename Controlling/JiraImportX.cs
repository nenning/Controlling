using System.Globalization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Controlling
{

    public class JiraData2
    {
        public string Type { get; init; }
        public string Key { get; init; }
        public string Summary { get; init; }
        public float Estimate { get; init; }
        public string[] SubTasks { get; init; }
        public override string ToString()
        {
            return $"Type: {Type}, Summary: {Summary}, Key: {Key}, Summary: {Summary}, Estimate: {Estimate}";
        }
    }
    internal class JiraImportX
    {

        string s = @"T	Key	Summary	Status	Resolution	Epic Link	Story Points	Assignee	Reporter	P	Created	Updated	Due	AT Points	Analysis Points	TestingPoints	Sprint	Components";
        string[] columnHeaders = new string[] {
            "T", "Key", "Summary", "Status", "Epic Link", "Story Points", "FE", "BE",
            "Assignee", "Reporter", "P", "Created", "Updated", "Due", "AT Points",
            "Analysis Points", "TestingPoints", "Sprint", "Components"
        };

        static string[] columns = { "Projekt-Nr.", "Mitarbeiter", "Datum", "Leistung", "Beschreibung", "V", "Abrechnung", "Anzahl", "Ansatz", "Betrag", "CR-Nr." };
        public static IEnumerable<TicketData> ParseExcelFile(string fileName)
        {
            var subtaskToParent = new Dictionary<string, string>();
            var jiraData = new List<TicketData>();
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
                    var firstCell = GetCellValue(cells[0], workbookPart);
                    if (string.IsNullOrEmpty(firstCell)) continue;
                    if (firstCell == "Total Stunden Mitarbeiter STD") break;
                    var data = new Booking
                    {
                        //ProjectNr = GetCellValue(cells[0], workbookPart),
                        Employee = GetCellValue(cells[1], workbookPart),
                        //Date = DateTime.ParseExact(GetCellValue(cells[2], workbookPart), "dd.MM.yyyy", CultureInfo.InvariantCulture),
                        Date = DateOnly.FromDateTime(DateTime.FromOADate(double.Parse(GetCellValue(cells[2], workbookPart)))),
                        WorkType = GetCellValue(cells[3], workbookPart),
                        WorkDescription = GetCellValue(cells[4], workbookPart),
                        Hours = float.Parse(GetCellValue(cells[7], workbookPart), CultureInfo.InvariantCulture),
                        Rate = float.Parse(GetCellValue(cells[8], workbookPart), CultureInfo.InvariantCulture),
                        TicketId = GetCellValue(cells[10], workbookPart)
                    };
                    //TODO jiraData.Add(data);
                    //for (int i = 0; i < cells.Length; i++)
                    //{
                    //    //string text = cells[i].InnerText;
                    //    string text = GetCellValue(cells[i], workbookPart);
                    //}
                }
            }
            return jiraData;
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

/*
 
                // Read the header row and add a new column named "Estimation"
                var headerRow = rows.First();
                var estimationColumnIndex = headerRow.Descendants<Cell>().Count() + 1;
                headerRow.AppendChild(new Cell
                {
                    CellValue = new CellValue("Estimation"),
                    DataType = CellValues.String,
                    CellReference = GetColumnName(estimationColumnIndex) + "1"
                });

                // Iterate through the data rows and add values to the "Estimation" column
                foreach (Row row in rows.Skip(1))
                {
                    Cell estimationCell = new Cell
                    {
                        CellValue = new CellValue("Sample Value"),
                        DataType = CellValues.String,
                        CellReference = GetColumnName(estimationColumnIndex) + row.RowIndex
                    };
                    row.AppendChild(estimationCell);
                }

                worksheetPart.Worksheet.Save();

 */