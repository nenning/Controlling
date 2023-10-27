using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Globalization;

namespace Controlling
{
    public class Contract
    {
        public Project Project { get; init; }
        public string Name { get; init; }
        public string Id { get; init; }
        public DateOnly StartDate { get; init; }
        public DateOnly EndDate { get; init; }
        public float Budget { get; init; }
        public float Estimate => Budget / 8.0f / 119.0f;
        public override string ToString()
        {
            return $"Contract {{ Project = {Project}, Name = {Name}, Id = {Id}, StartDate = {StartDate}, EndDate = {EndDate}, Budget = {Budget}, Estimate = {Estimate} }}";
        }
    }
    public class Project
    {
        public string Name { get; set; }
        public float DaysPerStoryPoint { get; set; }
        public string FilePrefix { get; set; }
        public string JiraKey { get; set; }
        public override string ToString()
        {
            return $"Project {{ Name = {Name}, JiraKey = {JiraKey}, DaysPerStoryPoint = {DaysPerStoryPoint}, FilePrefix = {FilePrefix} }}";
        }
    }
    public class Leftover
    {
        // columns: original ticket	follow-up tickets (comma-separated)
        public string OriginalTicketKey { get; set; }
        public IEnumerable<string> FollowUpTicketKeys { get; set; }
    }
    class ExcelFile : IDisposable {
        private bool disposedValue;
        private TemporaryFile file;
        private SpreadsheetDocument document;
        private WorkbookPart workbookPart;
        private Worksheet worksheet;
        private Sheets sheets;
        private Sheet sheet;

        public SharedStringTable SharedStringTable { get; private set; }

        public IEnumerable<Row> Rows { get { return worksheet.Descendants<Row>(); } }
        public ExcelFile(string filePath)
        {
            file = TemporaryFile.CreateCopy(filePath);
            filePath = file.FilePath;

            document = SpreadsheetDocument.Open(filePath, false);
            workbookPart = document.WorkbookPart;
            sheets = workbookPart.Workbook.GetFirstChild<Sheets>();
        }


        public void LoadSheet(string sheetName)
        {
            Sheet sheet = sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == sheetName);
            if (sheet == null)
            {
                throw new ArgumentException($"Sheet '{sheetName}' not found.");
            }
            worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;
            var sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            SharedStringTable = sstPart.SharedStringTable;
        }

        protected virtual void Dispose(bool disposing)
        {
            if (!disposedValue && disposing)
            {
                if (document != null)
                {
                    document.Dispose();
                }
                document = null;
                if (file != null) { 
                    file.Dispose(); 
                }
                file = null;
                disposedValue = true;
            }
        }

        public void Dispose()
        {
            Dispose(disposing: true);
            GC.SuppressFinalize(this);
        }
    }

    public class ProjectSettings
    {
        // contracts:  abacus id	project	name	startdate	enddate	Budget CHF	Budget days	Notes
        // projects: project	jira export file	SP to PT multiplier

        public IEnumerable<Contract> Contracts { get; private set; }
        public IEnumerable<Project> Projects { get; private set; }
        public List<Leftover> Leftovers { get; private set; }


        public ProjectSettings(string filePath)
        {
            using var excelFile = new ExcelFile(filePath);

            this.Projects = LoadProjects(excelFile);
            this.Contracts = LoadContracts(excelFile, this.Projects);
            this.Leftovers = LoadLeftovers(excelFile);

         }

        private static List<Project> LoadProjects(ExcelFile excelFile)
        {
            excelFile.LoadSheet("projects");
            var sst = excelFile.SharedStringTable;
            var projectList = new List<Project>();

            // Skip the first row (header)
            foreach (Row row in excelFile.Rows.Skip(1))
            {
                var project = new Project()
                {
                    Name = GetCellValue(row, 0, sst),
                    FilePrefix = GetCellValue(row, 1, sst),
                    JiraKey = GetCellValue(row, 2, sst),
                    DaysPerStoryPoint = float.Parse(GetCellValue(row, 3, sst)),
                };
                projectList.Add(project);
            }
            return projectList;
        }

        private static List<Contract> LoadContracts(ExcelFile excelFile, IEnumerable<Project> projects)
        {
            excelFile.LoadSheet("contracts");
            var sst = excelFile.SharedStringTable;
            var contractList = new List<Contract>();

            // Skip the first row (header)
            foreach (Row row in excelFile.Rows.Skip(1))
            {
                var contract = new Contract
                {
                    Id = GetCellValue(row, 0, sst),
                    Project = projects.FirstOrDefault(p => p.Name == GetCellValue(row, 1, sst)),
                    Name = GetCellValue(row, 2, sst),
                    StartDate = GetCellValue(row, 3, sst).ParseDate(),
                    EndDate = GetCellValue(row, 4, sst).ParseDate(),
                    Budget = float.Parse(GetCellValue(row, 5, sst), CultureInfo.InvariantCulture)
                };
                contractList.Add(contract);
            }
            return contractList;
        }

        private static List<Leftover> LoadLeftovers(ExcelFile excelFile)
        {
            excelFile.LoadSheet("leftovers");
            SharedStringTable sst = excelFile.SharedStringTable;
            var leftovers = new List<Leftover>();

            // Skip the first row (header)
            foreach (Row row in excelFile.Rows.Skip(1))
            {
                var leftover = new Leftover
                {
                    OriginalTicketKey = GetCellValue(row, 0, sst),
                    FollowUpTicketKeys = GetCellValue(row, 1, sst).Split(',')
                };

                leftovers.Add(leftover);
            }
            return leftovers;
        }



        private static string GetCellValue(Row row, int cellIndex, SharedStringTable sst)
        {
            Cell cell = row.Descendants<Cell>().ElementAt(cellIndex);

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                int ssid = int.Parse(cell.InnerText);
                return sst.ChildElements[ssid].InnerText;
            }
            else
            {
                return cell.InnerText;
            }
        }
    }

}
