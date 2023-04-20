// See https://aka.ms/new-console-template for more information
using Controlling;
using DocumentFormat.OpenXml.Wordprocessing;
using System.Configuration;


internal class Program
{
    private static void Main(string[] args)
    {
        var settings = new ProjectSettings(FindLatestFile("abacus projects"));

        // TODO could loop through reports by project.
        IEnumerable<TicketData> tasks = new List<TicketData>();
        Dictionary<string, string> subTasks = new();
        foreach (var project in settings.Projects) {
            var jiraData = JiraImporter.Import(FindLatestFile(project.FilePrefix), project);
            tasks = Enumerable.Concat(tasks, jiraData.Tasks);
            subTasks = jiraData.SubTasks.Concat(subTasks).ToDictionary(x => x.Key, x => x.Value);
        }
        var jiraImport = new JiraImport
        {
            Tasks = tasks,
            SubTasks = subTasks
        };

        var bookings = AbacusImport.ParseExcelFile(FindLatestFile("Leistungsauszug"), settings);

        foreach (var booking in  bookings)
        {
            var ticket = jiraImport.FindTask(booking.TicketId, booking.Contract.Project.JiraKey);
            if (ticket != null)
            {
                ticket.Hours += booking.Hours;
            } else if (ticket != null && !ticket.Key.EndsWith(booking.TicketId)) 
            {
                Console.WriteLine($"Wrong ticket entry: {booking}" );
            }
            // could improve class references
        }

        // TODO: could consider aggregated Tasks (leftovers)

        // Reports
        Reports.ShowLastBooking(bookings);
        Reports.ShowWrongEstimates(jiraImport.Tasks);
        Reports.ShowOutOfSprintBookings(bookings);
        Reports.ShowTicketsWithWorkWithoutEstimates(jiraImport.Tasks);
        Reports.ShowOBookingsByEmployeeBySprint(settings, bookings, jiraImport.Tasks);
        Console.ReadLine();

        static string GetWorkingDirectory()
        {
            return ConfigurationManager.AppSettings["workingDirectory"];
        }

        static string FindLatestFile(string filePrefix)
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

            return latestFile;
        }
    }
}