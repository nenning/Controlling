using DocumentFormat.OpenXml.Wordprocessing;
using System.Net.Sockets;

namespace Controlling
{
    public class Reports
    {
        private readonly IOutput output;
        private readonly ITimeProvider timeProvider;

        public Reports(IOutput output, ITimeProvider timeProvider)
        {
            this.output = output;
            this.timeProvider = timeProvider;
        }

        public void ShowLateBookings(IEnumerable<Booking> bookings)
        {
            output.WriteLine("----------------");
            output.WriteLine("Late bookings (>2d):");

            var employees = bookings.DistinctBy(a => a.Employee).Select(b => b.Employee).ToList();
            foreach (var employee in employees)
            {
                var latest = bookings.Where(b => b.Employee==employee).Aggregate((mostRecent, current) => current.Date > mostRecent.Date ? current : mostRecent);
                if (!latest.Date.IsMoreDaysAgoThan(21) && latest.Date.IsMoreDaysAgoThan(2)) {
                    output.WriteLine($" - {employee}: {latest.Date.DayMonth()} ({latest.Contract.Name})");
                }
            }
        }

        public void ShowWrongEstimates(IEnumerable<TicketData> tickets, ProjectSettings settings)
        {
            output.WriteLine("----------------");
            output.WriteLine("Wrong estimates:");
            // Idea: can handle leftovers here.
            
            foreach (var ticket in tickets.Where(t => t.TotalPoints > 0))
            {
                if (ticket.Percent > 1.1f && !ticket.Updated.IsMoreDaysAgoThan(16))
                {
                    output.WriteLine($" - {ticket.Key} actual: {ticket.Hours:N0}h, plan: {ticket.TotalPoints * 8 * ticket.Project.DaysPerStoryPoint}h ({ticket.Sprint}, {ticket.Status}, Updated: {ticket.Updated.DayMonth()}). {ticket.IssueType}: {ticket.Summary}");
                }
            }
        }

        public void ShowStoryEstimates(IEnumerable<TicketData> tickets, IEnumerable<Contract> contracts)
        {
            output.WriteLine("----------------");
            output.WriteLine("Story estimates:");
            foreach (var contract in contracts.Where(c => !c.EndDate.IsMoreDaysAgoThan(21))) {
                output.WriteLine($" - {contract.Name}:");
                foreach (var ticket in tickets.Where(t => (t.IssueType == "Story" || t.IssueType == "Task") && t.TotalPoints > 0 && t.Hours > 0.0f && t.Contract?.Id == contract.Id).OrderBy(t => t.Key))
                {
                    output.WriteLine($"  -- {ticket.Key} actual: {ticket.Hours:N0}h, plan: {ticket.TotalPoints * 8 * ticket.Project.DaysPerStoryPoint}h (Points: {ticket.TotalPointsText}, {ticket.Status}). {ticket.IssueType}: {ticket.Summary}");
                }
            }
        }

        public void ShowBookingsPerTicketType(IEnumerable<TicketData> tickets, IEnumerable<Contract> contracts)
        {
            output.WriteLine("----------------");
            output.WriteLine("Bookings per ticket type:");
            foreach (var contract in contracts.Where(c => !c.EndDate.IsMoreDaysAgoThan(32)))
            {
                var ticketsByType = tickets.Where(t => t.Hours > 0.0f && t.Contract?.Id == contract.Id).GroupBy(t => t.IssueType).OrderBy(t => t.Key);
                if (ticketsByType.Any())
                {
                    output.Write($" - {contract.Name}: ");
                    foreach (var group in ticketsByType)
                    {
                        double totalDays = group.Sum(t => t.Hours) / 8;
                        output.Write($"[{group.Key}: {totalDays:N1}d] ");
                    }
                    output.WriteLine();
                }
            }
        }


        public void ShowCostCeiling(IEnumerable<Booking> bookings)
        {
            output.WriteLine("----------------");
            output.WriteLine("Cost Ceiling:");
            // TODO: get rid of hardcoded values
            const double archCeiling = 400;
            const double devCeiling = 1750;
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
                    currentMonths = GetMonthsDifference(booking.Contract.StartDate, DateOnly.FromDateTime(timeProvider.Now));
                    totalMonths = GetMonthsDifference(booking.Contract.StartDate, booking.Contract.EndDate);
                }
            }
            output.WriteLine($" - Arch: {architectureTasks*100 / archCeiling:N2}% ({architectureTasks:N2}h of {archCeiling:N2}h)");
            output.WriteLine($" - Dev: {devTasks*100 / devCeiling:N2}% ({devTasks:N2}h of {devCeiling:N2}h)");
            output.WriteLine($" - Total: {(devTasks +architectureTasks) * 100 / (devCeiling+archCeiling):N0}% ({devTasks + architectureTasks:N0}h of {devCeiling + archCeiling:N0}h)");
            output.WriteLine($" - Time: {currentMonths*100 / totalMonths:N0}% ({currentMonths:N0} of {totalMonths:N0} months)");
        }

        public int GetMonthsDifference(DateOnly startDate, DateOnly endDate)
        {
            return (endDate.Year - startDate.Year) * 12 + endDate.Month - startDate.Month;
        }
        public void ShowMarginPerSprint(IEnumerable<TicketData> tickets, IEnumerable<Contract> contracts)
        {
            output.WriteLine("----------------");
            output.WriteLine("Estimation margins:");
            foreach (var contract in contracts.Where(c => !c.EndDate.IsMoreDaysAgoThan(32)))
            {
                double plan = 0; 
                double actual = 0;
                foreach (var ticket in tickets.Where(t => t.TotalPoints > 0 && t.Contract?.Id == contract.Id))
                {
                    plan += ticket.TotalPoints * 8 * ticket.Project.DaysPerStoryPoint;
                    actual += ticket.Hours;
                }
                output.WriteLine($" - {plan/actual:0.#} ({contract.Name}). Plan: {plan/8:0.#}d, Actual: {actual/8:0.#}d");
            }
        }

        public void ShowTicketsWithWorkWithoutEstimates(IEnumerable<TicketData> tickets)
        {
            output.WriteLine("----------------");
            output.WriteLine("No estimates:");

            foreach (var ticket in tickets.OrderBy(t => t.Sprint ?? string.Empty).Where(t => t.TotalPoints <= 0.000001f))
            {
                if (ticket.Hours > 1.0f && !ticket.Contract.EndDate.IsMoreDaysAgoThan(21))
                {
                    output.WriteLine($" - {ticket.Key} actual: {ticket.Hours:0.#}h ({ticket.Sprint}, {ticket.Status}, Updated: {ticket.Updated.DayMonth()}). {ticket.IssueType}: {ticket.Summary}");
                }
            }
        }

        public void ShowOutOfSprintBookings(IEnumerable<Booking> bookings)
        {
            output.WriteLine("----------------");
            output.WriteLine("Out-of-sprint bookings:");

            foreach (var booking in bookings)
            {
                if (booking.Date.IsMoreDaysAgoThan(21)) continue;
                if (booking.Date < booking.Contract.StartDate || booking.Date > booking.Contract.EndDate.AddDays(1))
                {
                    DateOnly startDate = booking.Contract.StartDate;
                    DateOnly endDate = booking.Contract.EndDate;
                    int diff = (booking.Date < startDate) ? startDate.DayNumber - booking.Date.DayNumber : booking.Date.DayNumber - endDate.DayNumber;
                    string ticketId = string.IsNullOrWhiteSpace(booking.TicketId) ? "-" : booking.TicketId;
                    output.WriteLine($" - {booking.Employee}: {booking.Hours}h on {booking.Date.DayMonth()} (Δ {diff}d). {booking.Contract.Name} ({booking.Contract.Id}; jira: {ticketId}): {startDate.DayMonth()}-{endDate.DayMonth()}");
                }
            }
        }

        public void ShowBookingsByEmployeeBySprint(IEnumerable<Contract> contracts, IEnumerable<Booking> bookings, IEnumerable<TicketData> tickets, IEnumerable<Person> persons)
        {
            output.WriteLine("----------------");
            output.WriteLine("Overview of bookings:");
            var employees = bookings.DistinctBy(a => a.Employee).Select(b => b.Employee).ToList();
            // Idea: could group by location as well and show distribution
            foreach (var contract in contracts.Where(c => !c.EndDate.IsMoreDaysAgoThan(21)))
            {
                output.WriteLine($"{contract.Name} ({contract.Id}): ");
                double workHours = 0;
                foreach (var employee in employees)
                {
                    double hours = bookings.Aggregate(0.0f, (total, next) => next.Employee==employee && next.Contract.Id == contract.Id ? total + next.Hours : total);
                    if (hours > 0)
                    {
                        output.WriteLine($" - {employee}: {Math.Round(hours / 8.0, 1)}d");
                    }
                    workHours += hours;
                }
                output.Write($" = Total: {workHours / 8.0:0.#}d, CHF {workHours * contract.HourlyRate:N0} (Plan: {contract.Budget / contract.HourlyRate / 8:0.#}d, CHF {contract.Budget:N0})");

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
                output.WriteLine($". Total cost: CHF {totalCost:N0}. Avg cost per hour: CHF {totalCost / totalHours:0.#}");
            }
        }

        public void ShowWarnings(IEnumerable<Contract> contracts, IEnumerable<Booking> bookings, IEnumerable<TicketData> tickets, IEnumerable<Person> persons)
        {
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
                output.WriteLine("----------------");

                output.WriteLine("Warning: Contract(s) without bookings found - check if they are part of the abacus export", isError: true);
                foreach (var contract in contractsWithoutBooking)
                {
                    output.WriteLine($" - {contract.Name} ({contract.Id})", isError: true);
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
                output.WriteLine($"Missing persons: {string.Join(", ", missingPersons)}", isError: true);
            }
        }

    }
}