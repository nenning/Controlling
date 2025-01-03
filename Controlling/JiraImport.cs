﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Controlling
{
    public class TicketData
    {
        public string IssueType { get; set; }
        public string Key { get; set; }
        public string Summary { get; set; }
        public double? StoryPoints { get; set; }
        public double? AnalysisPoints { get; set; }
        public double TotalPoints
        {
            get
            {
                if (Project.UseAnalysisPoints) return (StoryPoints ?? 0) + (AnalysisPoints ?? 0);
                else return StoryPoints ?? 0;
            }
        }
        public string TotalPointsText
        {
            get
            {
                if (Project.UseAnalysisPoints && TotalPoints > 0.000001)
                {
                    return $"{TotalPoints} (SP:{StoryPoints ?? 0}, AP:{AnalysisPoints ?? 0})";
                }
                else return $"{StoryPoints ?? 0}";
            }
        }
        public string Labels { get; set; }
        public string Assignee { get; set; }
        public string Reporter { get; set; }
        public string Priority { get; set; }
        public string Status { get; set; }
        public string Resolution { get; set; }
        public DateOnly Created { get; set; }
        public DateOnly Updated { get; set; }
        public DateOnly? DueDate { get; set; }
        public string Sprint { get; set; }
        public string EpicLink { get; set; }
        public string SubTasks { get; set; }
        public string Verrechenbarkeit { get; set; }
        public Project Project { get; set; }
        public Contract Contract { get; set; }

        // Calculated fields
        public double Hours { get; set; }
        public double Days { get {  return Hours/8.0; } }
        public double Percent
        {
            get { return (TotalPoints > 0.000001) ? Days / (TotalPoints * Project.DaysPerStoryPoint) : 0; }
        }
        public override string ToString()
        {
            return $"TicketData {{ IssueType = {IssueType}, Key = {Key}, Summary = {Summary}, TotalPoints = {TotalPointsText}, Status = {Status}, Sprint = {Sprint}, Days = {Days}, Percent = {Percent} }}";
        }

    }

    class JiraImport
    {
        public IEnumerable<TicketData> Tasks { get; init; }
        public IDictionary<string, string> SubTasks { get; init; }
        public TicketData FindTask(string key, string project)
        {
            if (string.IsNullOrEmpty(key) || key == "0") return null;
            if (!key.Contains("-")) { key = $"{project}-{key}".ToUpperInvariant(); }
            var result = Tasks.FirstOrDefault(t => t.Key == key);
            if (result == null)
            {
                if (SubTasks.TryGetValue(key, out key))
                {
                    result = Tasks.FirstOrDefault(t => t.Key == key);
                }
            }
            return result;
        }
    }
    
    class JiraImporter
    {
        public static JiraImport Import(string filePath, Project project)
        {
            return ReadExcelFile(filePath, project);
        }
        private static string GetColumnName(Dictionary<string, int> columnIndexMap, Cell cell)
        {
            int cellColumnIndex = GetColumnIndex(cell.CellReference.Value);
            return columnIndexMap.FirstOrDefault(x => x.Value == cellColumnIndex).Key;
        }

        private static int GetColumnIndex(string cellReference)
        {
            int columnIndex = 0;
            int multiplier = 1;

            for (int i = cellReference.Length - 1; i >= 0; i--)
            {
                if (char.IsLetter(cellReference[i]))
                {
                    columnIndex += (cellReference[i] - 'A' + 1) * multiplier;
                    multiplier *= 26;
                }
            }

            return columnIndex;
        }
        private static JiraImport ReadExcelFile(string filePath, Project project)
        {
            using var file = TemporaryFile.CreateCopy(filePath);
            filePath = file.FilePath;

            var jiraDataList = new List<TicketData>();
            var subTaskParents = new Dictionary<string, string>();

            using (SpreadsheetDocument doc = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = doc.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().LastOrDefault();
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);
                SharedStringTablePart stringTablePart = workbookPart.GetPartsOfType<SharedStringTablePart>().FirstOrDefault();
                var rows = worksheetPart.Worksheet.Descendants<Row>();

                // Read the header row and map column names to indices
                var headerRow = rows.First();
                var columnIndexMap = new Dictionary<string, int>();
                int columnIndex = 1;
                foreach (Cell cell in headerRow.Descendants<Cell>())
                {
                    string columnName = cell.GetCellValue(stringTablePart);
                    columnIndexMap[columnName] = columnIndex++;
                }
                try
                {
                    // Read the data rows
                    foreach (Row row in rows.Skip(1))
                    {
                        var rowData = new TicketData();
                        rowData.Project = project;
                        bool skip = false;
                        foreach (Cell cell in row.Descendants<Cell>())
                        {
                            if (cell.CellFormula != null)
                            {
                                skip = true;
                                break;
                            }
                            string columnName = GetColumnName(columnIndexMap, cell);
                            string cellValue = cell.GetCellValue(stringTablePart);

                            SetPropertyValue(rowData, columnName, cellValue);
                        }
                        // Idea: could handle multiple sprints: "Team AIG - Sprint 30, Sofia - Sprint 36". Take latest only and only sprints in the list
                        if (!skip && rowData.IssueType.ToLowerInvariant() != "sub-task")
                        {
                            jiraDataList.Add(rowData);
                            if (!string.IsNullOrEmpty(rowData.SubTasks))
                            {
                                var subTasks = rowData.SubTasks.Split(new string[] { ",", ";" }, StringSplitOptions.TrimEntries | StringSplitOptions.RemoveEmptyEntries);
                                foreach (var subTask in subTasks)
                                {
                                    if (!subTaskParents.ContainsKey(subTask)) subTaskParents[subTask] = rowData.Key;
                                }
                            }
                        }                        
                    }
                } catch (System.FormatException) { }
            }

            var result = new JiraImport
            {
                Tasks = jiraDataList,
                SubTasks = subTaskParents
            };
            return result;
        }

        // LGS: Issue Type	Key	Summary	Story Points	Labels	Assignee	Reporter	Priority	Status	Resolution	Created	Updated	Due date	Sprint	Epic Link	Sub-tasks	Verrechenbarkeit
        // ISR: T	Key	Summary	Status	Resolution	Feature Link	Story Points	Assignee	Reporter	P	Created	Updated	Due	AT Points	Analysis Points	TestingPoints	Sprint	Components	Sub-Tasks
        private static void SetPropertyValue(TicketData rowData, string columnName, string cellValue)
        {
            switch (columnName)
            {
                case "Issue Type" or "T":
                    rowData.IssueType = cellValue;
                    break;
                case "Key":
                    rowData.Key = cellValue;
                    break;
                case "Summary":
                    rowData.Summary = cellValue;
                    break;
                case "Story Points":
                    rowData.StoryPoints = string.IsNullOrWhiteSpace(cellValue) ? (double?)null : double.Parse(cellValue);
                    break;
                case "Analysis Points":
                    rowData.AnalysisPoints = string.IsNullOrWhiteSpace(cellValue) ? (double?)null : double.Parse(cellValue);
                    break;
                case "Labels":
                    rowData.Labels = cellValue;
                    break;
                case "Assignee":
                    rowData.Assignee = cellValue;
                    break;
                case "Reporter":
                    rowData.Reporter = cellValue;
                    break;
                case "Priority" or "P":
                    rowData.Priority = cellValue;
                    break;
                case "Status":
                    rowData.Status = cellValue;
                    break;
                case "Resolution":
                    rowData.Resolution = cellValue;
                    break;
                case "Created":
                    rowData.Created = cellValue.ParseDate();
                    break;
                case "Updated":
                    rowData.Updated = cellValue.ParseDate();
                    break;
                case "Due date" or "Due":
                    rowData.DueDate = string.IsNullOrWhiteSpace(cellValue) ? (DateOnly?)null : cellValue.ParseDate();
                    break;
                case "Sprint":
                    rowData.Sprint = cellValue;
                    break;
                case "Epic Link" or "Feature Link":
                    rowData.EpicLink = cellValue;
                    break;
                case "Sub-tasks" or "Sub-Tasks":
                    rowData.SubTasks = cellValue;
                    break;
                case "Verrechenbarkeit":
                    rowData.Verrechenbarkeit = cellValue;
                    break;
                default:
                    break;
            }
        }


    }
}
