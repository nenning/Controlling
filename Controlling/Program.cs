// See https://aka.ms/new-console-template for more information
using Controlling;
using System.Configuration;
using System.Diagnostics;

/// <summary>
/// Use app.config.local to override the app.config settings
/// </summary>
internal class Program
{
    private static readonly ITimeProvider timeProvider = new TimeProvider();
    private static readonly IOutput output = new Output();

    private static void Main(string[] args)
    {
        Console.OutputEncoding = System.Text.Encoding.UTF8;
        try
        {
            string workingDirectory = ConfigurationManager.AppSettings["App:WorkingDirectory"];
            string projectsFilePrefix = ConfigurationManager.AppSettings["App:ProjectsFilePrefix"];
            string bookingsFilePrefix = ConfigurationManager.AppSettings["App:BookingsFilePrefix"];

            output.WriteLine($"Starting import from {workingDirectory}...");
            TerminateOldProcesses();

            Console.Title = "Waiting for OneDrive Sync...";
            ForceOneDriveSync(workingDirectory);

            ProjectManager projectManager = new (workingDirectory, projectsFilePrefix, bookingsFilePrefix, output, timeProvider);
            projectManager.DoControlling();

            output.WriteLine();
        }
        catch (Exception ex)
        {
            output.WriteLine(ex.ToString());
        }
        output.WriteLine("Press any key...");
        Console.ReadLine();
    }
    static void TerminateOldProcesses()
    {
        if (Debugger.IsAttached) return;
        var currentProcess = Process.GetCurrentProcess();

        Process[] processes = Process.GetProcessesByName(currentProcess.ProcessName);
        processes = processes.Where(p => p.Id != currentProcess.Id && p.MainModule.FileName == currentProcess.MainModule.FileName)
            .OrderBy(p => p.StartTime)
            .ToArray();

        foreach (Process process in processes)
        {
            try
            {
                process.Kill();
            }
            catch (Exception)
            {
                // ignore
            }
        }
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
}