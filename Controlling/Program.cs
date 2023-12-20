// See https://aka.ms/new-console-template for more information
using Controlling;
using System.Configuration;
using System.Diagnostics;

internal class Program
{
    private static void Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;

        try
        {
            /*
            // Get the Jira configuration values
            string baseUrl = ConfigurationManager.AppSettings["Jira:BaseUrl"];
            string username = ConfigurationManager.AppSettings["Jira:UserName"];
            string apiKey = ConfigurationManager.AppSettings["Jira:ApiKey"];
            string jql = ConfigurationManager.AppSettings["Jira:Jql"];

            var jiraClient = new JiraApiClient(baseUrl, username, apiKey);
            var issues = jiraClient.GetAllIssuesWithFields(jql);
            */
            var workingDirectory = GetWorkingDirectory();
            Console.WriteLine($"Starting import from {workingDirectory}...");

            Console.Title = "Waiting for OneDrive Sync...";
            ForceOneDriveSync(workingDirectory);

            Console.Title = "Importing files...";
            var settings = new ProjectSettings(FindLatestFile("abacus projects"));

            IEnumerable<TicketData> tasks = new List<TicketData>();
            Dictionary<string, string> subTasks = new();
            foreach (var project in settings.Projects)
            {
                if (project.JiraKey == "undefined") continue;
                var jiraData = JiraImporter.Import(FindLatestFile(project.FilePrefix), project);
                tasks = Enumerable.Concat(tasks, jiraData.Tasks);
                subTasks = jiraData.SubTasks.Concat(subTasks).ToDictionary(x => x.Key, x => x.Value);
            }
            var jiraImport = new JiraImport
            {
                Tasks = tasks,
                SubTasks = subTasks
            };

            var bookings = AbacusImport.ParseExcelFile(FindLatestFile("Leistungsauszug", warnAboutOldFile:true), settings);
            foreach (var booking in bookings)
            {
                var ticket = jiraImport.FindTask(booking.TicketId, booking.Contract.Project.JiraKey);
                if (ticket != null)
                {
                    ticket.Hours += booking.Hours;
                    booking.TicketId = ticket.Key;
                    ticket.Contract = booking.Contract;
                }
                else if (ticket != null && !ticket.Key.EndsWith(booking.TicketId))
                {
                    Tools.UseErrorColors();
                    Console.WriteLine($"Wrong ticket entry: {booking}");
                    Tools.UseStandardColors();
                }
                // could improve class references
            }

            // Idea: could consider aggregated Tasks (leftovers)
            Console.Title = "Controlling";
            ShowReports(settings, jiraImport, bookings);

            Console.WriteLine();
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.ToString());
        }
        Console.WriteLine("Press any key...");
        Console.ReadLine();
    }
        
    private static void ForceOneDriveSync(string folderPath)
    {
        if (!Directory.Exists(folderPath))
        {
            throw new InvalidOperationException("Folder not found {folderPath}.");
        }
        string tempFilePath = Path.Combine(folderPath, "temp_sync_trigger.txt");
        try
        {
            // Create a temporary file in the folder
            File.WriteAllText(tempFilePath, "OneDrive sync trigger");
            // Wait for a short time to allow OneDrive to detect the change
            var delayMilliseconds = Debugger.IsAttached ? 0 : 5000;
            Thread.Sleep(delayMilliseconds);
        }
        finally
        {
            File.Delete(tempFilePath);
        }
    }
    private static void ShowReports(ProjectSettings settings, JiraImport jiraImport, IEnumerable<Booking> bookings)
    {
        foreach (var project in settings.Projects)
        {
            PrintProjectTitle(project);

            var currentBookings = bookings.Where(x => x.Contract.Project.Name == project.Name).ToList();
            if (project.JiraKey == "undefined")
            {
                Reports.ShowCostCeiling(currentBookings);
                continue;
            }
            var currentTickets = jiraImport.Tasks.Where(x => x.Project.Name == project.Name).ToList();
            var currentContracts = settings.Contracts.Where(x => x.Project.Name == project.Name).ToList();

            Reports.ShowWarnings(currentContracts, currentBookings, currentTickets, settings.Persons);
            Reports.ShowLateBooking(currentBookings);
            Reports.ShowWrongEstimates(currentTickets, settings);
            Reports.ShowOutOfSprintBookings(currentBookings);
            Reports.ShowTicketsWithWorkWithoutEstimates(currentTickets);
            Reports.ShowBookingsByEmployeeBySprint(currentContracts, currentBookings, currentTickets, settings.Persons);
            Reports.ShowStoryEstimates(currentTickets, currentContracts);
            Reports.ShowMarginPerSprint(currentTickets, currentContracts);
        }
    }

    private static void PrintProjectTitle(Project project)
    {
        var projectText = $"Project: {project.Name}";
        var line = new string('*', projectText.Length + 6);
        Console.WriteLine();
        Console.WriteLine();
        Console.WriteLine(line);
        Console.Write("*  ");
        Console.BackgroundColor = ConsoleColor.DarkBlue;
        Console.ForegroundColor = ConsoleColor.White;
        Console.Write(projectText);
        Console.ResetColor();
        Console.WriteLine("  *");
        Console.WriteLine(line);
    }

    static string GetWorkingDirectory()
    {
        return ConfigurationManager.AppSettings["App:WorkingDirectory"];
    }

    static string FindLatestFile(string filePrefix, bool warnAboutOldFile = false)
    {
        string directoryPath = GetWorkingDirectory();
        string searchPattern = filePrefix + "*";
        string latestFile = null;
        DateTime latestDate = DateTime.MinValue;
        var foundFiles = new List<string>();

        DirectoryInfo directory = new DirectoryInfo(directoryPath);

        foreach (FileInfo file in directory.GetFiles(searchPattern))
        {
            foundFiles.Add(file.FullName);
            if (file.CreationTime > latestDate)
            {
                latestFile = file.FullName;
                latestDate = file.CreationTime;
            }
        }
        // clean up old files
        foreach (var file in foundFiles)
        {
            if (file != latestFile)
                File.Delete(file);
        }
        var created = File.GetLastWriteTime(latestFile);
        Console.WriteLine($" - importing: {Path.GetFileName(latestFile)} (updated: {created})");
        if (warnAboutOldFile && DateTime.Now.Subtract(TimeSpan.FromDays(3)) > created)
        {
            Tools.UseErrorColors();
            Console.WriteLine("     Outdated!     ");
            Tools.UseStandardColors();
        }


        return latestFile;
    }
}