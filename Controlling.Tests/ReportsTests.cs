using Xunit;
using Controlling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Controlling.Tests.util;
using System.Reflection;
using Microsoft.VisualStudio.TestPlatform.Utilities;

namespace Controlling.Tests
{
    public class ReportsTests
    {

        [Fact()]
        public void LoadProjectsTest()
        {
            string workingDirectory = this.GetType().Assembly.Location;
            string projectsFile = Path.Combine(workingDirectory, "abacus projects.xlsx");
            string bookingsFile = Path.Combine(workingDirectory, "Leistungsauszug.xlsx");

            var settings = new ProjectSettings(projectsFile);

            IEnumerable<TicketData> tasks = new List<TicketData>();
            Dictionary<string, string> subTasks = new();
        }

        [Fact()]
        public void ShowLateBookingTest()
        {
            // Arrange
            FakeOutput output = new();
            output.SetExpectation("Late booking: 2023-03-03");
            FakeTimeProvider timeProvider = new();
            timeProvider.PresetNow(new DateTime(2023, 3, 3));

            string workingDirectory = this.GetType().Assembly.Location;
            string projectsFile = Path.Combine(workingDirectory, "abacus projects.xlsx");
            string bookingsFile = Path.Combine(workingDirectory, "Leistungsauszug.xlsx");

            var settings = new ProjectSettings(projectsFile);

            IEnumerable<TicketData> tasks = new List<TicketData>();
            Dictionary<string, string> subTasks = new();
            foreach (var project in settings.Projects)
            {
                var jiraData = JiraImporter.Import(project.FilePrefix, project);
                tasks = Enumerable.Concat(tasks, jiraData.Tasks);
                subTasks = jiraData.SubTasks.Concat(subTasks).ToDictionary(x => x.Key, x => x.Value);
            }
            var jiraImport = new JiraImport
            {
                Tasks = tasks,
                SubTasks = subTasks
            };

            var bookings = AbacusImport.ParseExcelFile(bookingsFile, settings, output);
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

            // Act
            var reports = new Reports(output, timeProvider);
            reports.ShowLateBookings(bookings);
            // Assert
            output.Verify();
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

        [Fact()]
        public void ShowWrongEstimatesTest()
        {
            Assert.True(false, "This test needs an implementation");
        }

        [Fact()]
        public void ShowStoryEstimatesTest()
        {
            Assert.True(false, "This test needs an implementation");
        }

        [Fact()]
        public void ShowCostCeilingTest()
        {
            Assert.True(false, "This test needs an implementation");
        }

        [Fact()]
        public void GetMonthsDifferenceTest()
        {
            Assert.True(false, "This test needs an implementation");
        }

        [Fact()]
        public void ShowMarginPerSprintTest()
        {
            Assert.True(false, "This test needs an implementation");
        }

        [Fact()]
        public void ShowTicketsWithWorkWithoutEstimatesTest()
        {
            Assert.True(false, "This test needs an implementation");
        }

        [Fact()]
        public void ShowOutOfSprintBookingsTest()
        {
            Assert.True(false, "This test needs an implementation");
        }

        [Fact()]
        public void ShowBookingsByEmployeeBySprintTest()
        {
            Assert.True(false, "This test needs an implementation");
        }

        [Fact()]
        public void ShowWarningsTest()
        {
            Assert.True(false, "This test needs an implementation");
        }
    }
}