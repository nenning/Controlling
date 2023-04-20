using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

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
    public class ProjectSettings
    {
        // contracts:  abacus id	project	name	startdate	enddate	Budget CHF	Budget days	Notes
        // projects: project	jira export file	SP to PT multiplier

        public IEnumerable<Contract> Contracts { get; private set; }
        public IEnumerable<Project> Projects { get; private set; }


        public ProjectSettings(string filePath)
        {
            using SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false);
            WorkbookPart workbookPart = doc.WorkbookPart;
            Sheets sheets = workbookPart.Workbook.GetFirstChild<Sheets>();

            this.Projects = LoadProjects(workbookPart, sheets);
            this.Contracts = LoadContracts(workbookPart, sheets, this.Projects);

         }

        // TODO
        // original ticket	follow-up tickets (comma-separated)
        // ISRTOZEM-392	ISRTOZEM-534

        private static List<Project> LoadProjects(WorkbookPart workbookPart, Sheets sheets)
        {
            var sheet = sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == "projects");
            if (sheet == null)
            {
                throw new ArgumentException("Sheet 'projects' not found.");
            }

            var worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;
            var sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            var sst = sstPart.SharedStringTable;
            var projectList = new List<Project>();

            // Skip the first row (header)
            foreach (Row row in worksheet.Descendants<Row>().Skip(1))
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

        private static List<Contract> LoadContracts(WorkbookPart workbookPart, Sheets sheets, IEnumerable<Project> projects)
        {
            var sheet = sheets.Elements<Sheet>().FirstOrDefault(s => s.Name == "contracts");
            if (sheet == null)
            {
                throw new ArgumentException("Sheet 'contracts' not found.");
            }

            var worksheet = ((WorksheetPart)workbookPart.GetPartById(sheet.Id)).Worksheet;
            var sstPart = workbookPart.GetPartsOfType<SharedStringTablePart>().First();
            var sst = sstPart.SharedStringTable;
            var contractList = new List<Contract>();

            // Skip the first row (header)
            foreach (Row row in worksheet.Descendants<Row>().Skip(1))
            {
                var contract = new Contract
                {
                    Id = GetCellValue(row, 0, sst),
                    Project = projects.FirstOrDefault (p => p.Name == GetCellValue(row, 1, sst)),
                    Name = GetCellValue(row, 2, sst),
                    StartDate = ExcelHelper.ParseDate(GetCellValue(row, 3, sst)),
                    EndDate = ExcelHelper.ParseDate(GetCellValue(row, 4, sst)),
                    Budget = float.Parse(GetCellValue(row, 5, sst), CultureInfo.InvariantCulture)
                };
                foreach (var project in projects)
                {

                }
                contractList.Add(contract);
            }
            return contractList;
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
