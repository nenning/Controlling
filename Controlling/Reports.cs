﻿using DocumentFormat.OpenXml.Wordprocessing;
using System.Net.Sockets;

namespace Controlling
{
    public class Reports
    {

        public static void ShowLateBooking(IEnumerable<Booking> bookings)
        {
            Console.WriteLine("----------------");
            Console.WriteLine("Late bookings (>2d):");

            var employees = bookings.DistinctBy(a => a.Employee).Select(b => b.Employee).ToList();
            foreach (var employee in employees)
            {
                var latest = bookings.Where(b => b.Employee==employee).Aggregate((mostRecent, current) => current.Date > mostRecent.Date ? current : mostRecent);
                if (!latest.Date.IsMoreDaysAgoThan(21) && latest.Date.IsMoreDaysAgoThan(2)) {
                    Console.WriteLine($" - {employee}: {latest.Date.DayMonth()} ({latest.Contract.Name})");
                }
            }
        }

        public static void ShowWrongEstimates(IEnumerable<TicketData> tickets, ProjectSettings settings)
        {
            Console.WriteLine("----------------");
            Console.WriteLine("Wrong estimates:");
            // Idea: can handle leftovers here.
            
            foreach (var ticket in tickets.Where(t => t.StoryPoints.HasValue))
            {
                if (ticket.Percent > 1.1f && !ticket.Updated.IsMoreDaysAgoThan(16))
                {
                    Console.WriteLine($" - {ticket.Key} actual: {ticket.Hours:N0}h, plan: {ticket.StoryPoints * 8 * ticket.Project.DaysPerStoryPoint}h ({ticket.Sprint}, Status: {ticket.Status}, Updated: {ticket.Updated.DayMonth()}). {ticket.IssueType}: {ticket.Summary}");
                }
            }
        }

        public static void ShowStoryEstimates(IEnumerable<TicketData> tickets, IEnumerable<Contract> contracts)
        {
            Console.WriteLine("----------------");
            Console.WriteLine("Story estimates:");
            foreach (var contract in contracts.Where(c => !c.EndDate.IsMoreDaysAgoThan(21))) {
                Console.WriteLine($" - {contract.Name}:");
                foreach (var ticket in tickets.Where(t => (t.IssueType == "Story" || t.IssueType == "Task") && t.StoryPoints.HasValue && t.Hours > 0.0f && t.Contract?.Id == contract.Id).OrderBy(t => t.Key))
                {
                    Console.WriteLine($"  -- {ticket.Key} actual: {ticket.Hours:N0}h, plan: {ticket.StoryPoints * 8 * ticket.Project.DaysPerStoryPoint}h (Status: {ticket.Status}). {ticket.IssueType}: {ticket.Summary}");
                }
            }
        }

        
        public static void ShowCostCeiling(IEnumerable<Booking> bookings)
        {
            Console.WriteLine("----------------");
            Console.WriteLine("Cost Ceiling:");
            const double archCeiling = 400;
            const double devCeiling = 950;
            double architectureTasks = 0;
            double devTasks = 0;
            int currentMonths = 0;
            int totalMonths = 0;
            foreach (var booking in bookings)
            {
                if (booking.WorkType == "fachliche Konzeption")
                {
                    architectureTasks += booking.Hours;
                } else
                {
                    devTasks += booking.Hours;
                }
                if (currentMonths == 0 && totalMonths == 0) {
                    currentMonths = GetMonthsDifference(booking.Contract.StartDate, DateOnly.FromDateTime(DateTime.Now));
                    totalMonths = GetMonthsDifference(booking.Contract.StartDate, booking.Contract.EndDate);
                }
            }
            Console.WriteLine($" - Arch: {architectureTasks*100 / archCeiling:N0}% ({architectureTasks:N0}h of {archCeiling:N0}h)");
            Console.WriteLine($" - Dev: {devTasks*100 / devCeiling:N0}% ({devTasks:N0}h of {devCeiling:N0}h)");
            Console.WriteLine($" - Total: {(devTasks +architectureTasks) * 100 / (devCeiling+archCeiling):N0}% ({devTasks + architectureTasks:N0}h of {devCeiling + archCeiling:N0}h)");
            Console.WriteLine($" - Time: {currentMonths*100 / totalMonths:N0}% ({currentMonths:N0} of {totalMonths:N0} months)");
        }

        public static int GetMonthsDifference(DateOnly startDate, DateOnly endDate)
        {
            return (endDate.Year - startDate.Year) * 12 + endDate.Month - startDate.Month;
        }
        public static void ShowMarginPerSprint(IEnumerable<TicketData> tickets, IEnumerable<Contract> contracts)
        {
            Console.WriteLine("----------------");
            Console.WriteLine("Margins per Sprint:");
            foreach (var contract in contracts)
            {
                double plan = 0; 
                double actual = 0;
                foreach (var ticket in tickets.Where(t => t.IssueType == "Story" && t.StoryPoints.HasValue && t.Hours > 0.0f && t.Contract?.Id == contract.Id))
                {
                    plan += ticket.StoryPoints.Value * 8 * ticket.Project.DaysPerStoryPoint;
                    actual += ticket.Hours;
                }
                Console.WriteLine($" - {plan/actual:N2} ({contract.Name}). Plan: {plan/8:N2}, Actual: {actual/8:N2}");
            }
        }

        public static void ShowTicketsWithWorkWithoutEstimates(IEnumerable<TicketData> tickets)
        {
            Console.WriteLine("----------------");
            Console.WriteLine("No estimates:");

            foreach (var ticket in tickets.Where(t => !t.StoryPoints.HasValue || t.StoryPoints.Value==0))
            {
                if (ticket.Hours > 1.0f && !ticket.Contract.EndDate.IsMoreDaysAgoThan(21))
                {
                    Console.WriteLine($" - {ticket.Key} actual: {ticket.Hours}h ({ticket.Sprint}, Status: {ticket.Status}, Updated: {ticket.Updated.DayMonth()}). {ticket.IssueType}: {ticket.Summary}");
                }
            }
        }

        public static void ShowOutOfSprintBookings(IEnumerable<Booking> bookings)
        {
            Console.WriteLine("----------------");
            Console.WriteLine("Out-of-sprint bookings:");

            foreach (var booking in bookings)
            {
                if (booking.Date.IsMoreDaysAgoThan(21)) continue;
                if (booking.Date < booking.Contract.StartDate || booking.Date > booking.Contract.EndDate.AddDays(1))
                {
                    DateOnly startDate = booking.Contract.StartDate;
                    DateOnly endDate = booking.Contract.EndDate;
                    int diff = (booking.Date < startDate) ? startDate.DayNumber - booking.Date.DayNumber : booking.Date.DayNumber - endDate.DayNumber;
                    string ticketId = string.IsNullOrWhiteSpace(booking.TicketId) ? "-" : booking.TicketId;
                    Console.WriteLine($" - {booking.Employee}: {booking.Hours}h on {booking.Date.DayMonth()} (Δ {diff}d). {booking.Contract.Name} ({booking.Contract.Id}; jira: {ticketId}): {startDate.DayMonth()}-{endDate.DayMonth()}");
                }
            }
        }

        public static void ShowBookingsByEmployeeBySprint(IEnumerable<Contract> contracts, IEnumerable<Booking> bookings, IEnumerable<TicketData> tickets, IEnumerable<Person> persons)
        {
            Console.WriteLine("----------------");
            Console.WriteLine("Overview of bookings per sprint:");
            var employees = bookings.DistinctBy(a => a.Employee).Select(b => b.Employee).ToList();
            // Idea: could group by location as well and show distribution
            foreach (var contract in contracts.Where(c => !c.EndDate.IsMoreDaysAgoThan(21)))
            {
                Console.WriteLine($"{contract.Name} ({contract.Id}): ");
                double workHours = 0;
                foreach (var employee in employees)
                {
                    double hours = bookings.Aggregate(0.0f, (total, next) => next.Employee==employee && next.Contract.Id == contract.Id ? total + next.Hours : total);
                    if (hours > 0)
                    {
                        Console.WriteLine($" - {employee}: {Math.Round(hours / 8.0, 1)}d");
                    }
                    workHours += hours;
                }
                Console.Write($" = Total: {Math.Round(workHours / 8.0, 1)}d, {workHours * 119:N0} CHF (Plan: {contract.Budget / 119 / 8}d, {contract.Budget:N0} CHF)");

                double totalHours = 0.0;
                double totalCost = 0.0;
                foreach (var booking in bookings.Where(b => b.Contract == contract))
                {
                    var person = persons.FirstOrDefault(p => p.Name == booking.Employee);
                    if (person?.HourlyRate > 1.0f)
                    {
                        totalHours += booking.Hours;
                        totalCost += booking.Hours * person.HourlyRate;
                    }
                }
                Console.WriteLine($". Blended rate per hour: CHF {totalCost / totalHours:N2}");
            }
        }

        public static void ShowWarnings(IEnumerable<Contract> contracts, IEnumerable<Booking> bookings, IEnumerable<TicketData> tickets, IEnumerable<Person> persons)
        {
            Tools.UseErrorColors();

            var contractsWithoutBooking = new List<Contract>();
            foreach (var contract in contracts)
            {
                if (!bookings.Any(b=>b.Contract.Id == contract.Id))
                {
                    contractsWithoutBooking.Add(contract);
                }
            }
            if (contractsWithoutBooking.Count > 0)
            {
                Console.WriteLine("----------------");

                Console.WriteLine("Warning: Contract(s) without bookings found - check if they are part of the abacus export");
                foreach (var contract in contractsWithoutBooking)
                {
                    Console.WriteLine($" - {contract.Name} ({contract.Id})");
                }
            }
            
            var missingPersons = new HashSet<string>();
            foreach (var booking in bookings)
            {
                if (!persons.Any(p => p.Name == booking.Employee))
                {
                    missingPersons.Add(booking.Employee);
                }
            }
            if (missingPersons.Count > 0)
            {
                Console.WriteLine($"Missing persons: {string.Join(", ", missingPersons)}");
            }
            Tools.UseStandardColors();
        }

    }
}