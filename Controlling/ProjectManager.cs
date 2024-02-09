// See https://aka.ms/new-console-template for more information
using System;

namespace Controlling
{
    public class ProjectManager
    {
        private readonly string workingDirectory;
        private readonly string projectsFilePrefix;
        private readonly string bookingsFilePrefix;
        private readonly IOutput output;
        private readonly ITimeProvider timeProvider;

        public ProjectManager(string workingDirectory, string projectsFilePrefix, string bookingsFilePrefix, IOutput output, ITimeProvider timeProvider)
        {
            this.workingDirectory = workingDirectory;
            this.projectsFilePrefix = projectsFilePrefix;
            this.bookingsFilePrefix = bookingsFilePrefix;
            this.output = output;
            this.timeProvider = timeProvider;
        }
        public void DoControlling()
        {
            Console.Title = "Importing files...";
            var settings = new ProjectSettings(FindLatestFile(projectsFilePrefix));

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

            var bookings = AbacusImport.ParseExcelFile(FindLatestFile(bookingsFilePrefix, warnAboutOldFile: true), settings, output);
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
                    output.WriteLine($"Wrong ticket entry: {booking}", isError: true);
                }
                // could improve class references
            }

            // Idea: could consider aggregated Tasks (leftovers)
            Console.Title = "Controlling";
            ShowReports(settings, jiraImport, bookings);
        }

        private void ShowReports(ProjectSettings settings, JiraImport jiraImport, IEnumerable<Booking> bookings)
        {
            var reports = new Reports(output, timeProvider);
            foreach (var project in settings.Projects)
            {
                PrintProjectTitle(project);

                var currentBookings = bookings.Where(x => x.Contract.Project.Name == project.Name).ToList();
                if (project.JiraKey == "undefined")
                {
                    reports.ShowLateBookings(currentBookings);
                    reports.ShowCostCeiling(currentBookings);
                    continue;
                }
                var currentTickets = jiraImport.Tasks.Where(x => x.Project.Name == project.Name).ToList();
                var currentContracts = settings.Contracts.Where(x => x.Project.Name == project.Name).ToList();

                reports.ShowWarnings(currentContracts, currentBookings, currentTickets, settings.Persons);
                reports.ShowLateBookings(currentBookings);
                reports.ShowWrongEstimates(currentTickets, settings);
                reports.ShowOutOfSprintBookings(currentBookings);
                reports.ShowTicketsWithWorkWithoutEstimates(currentTickets);
                reports.ShowBookingsByEmployeeBySprint(currentContracts, currentBookings, currentTickets, settings.Persons);
                reports.ShowStoryEstimates(currentTickets, currentContracts);
                reports.ShowMarginPerSprint(currentTickets, currentContracts);
            }
        }
        private void PrintProjectTitle(Project project)
        {
            var projectText = $"Project: {project.Name}";
            var line = new string('*', projectText.Length + 6);
            output.WriteLine();
            output.WriteLine();
            output.WriteLine(line);
            output.Write("*  ");
            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.ForegroundColor = ConsoleColor.White;
            output.Write(projectText);
            Console.ResetColor();
            output.WriteLine("  *");
            output.WriteLine(line);
        }
        private string FindLatestFile(string filePrefix, bool warnAboutOldFile = false)
        {
            string searchPattern = filePrefix + "*";
            string latestFile = null;
            DateTime latestDate = DateTime.MinValue;
            var foundFiles = new List<string>();

            DirectoryInfo directory = new DirectoryInfo(workingDirectory);

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
            output.WriteLine($" - importing: {Path.GetFileName(latestFile)} (updated: {created})");
            if (warnAboutOldFile && timeProvider.Now.Subtract(TimeSpan.FromDays(3)) > created)
            {
                output.WriteLine("     Outdated!     ", isError: true);
            }


            return latestFile;
        }
    }
}