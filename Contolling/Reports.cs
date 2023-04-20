using DocumentFormat.OpenXml.Vml;

namespace Controlling
{
    public class Reports
    {

        public static void ShowLastBooking(IEnumerable<Booking> bookings)
        {
            Console.WriteLine("----------------");
            Console.WriteLine("Latest bookings:");

            var employees = bookings.DistinctBy(a => a.Employee).Select(b => b.Employee).ToList();
            foreach (var employee in employees)
            {
                var latest = bookings.Where(b => b.Employee==employee).Aggregate((mostRecent, current) => current.Date > mostRecent.Date ? current : mostRecent);
                if (!latest.Date.IsMoreDaysAgoThan(21)) {
                    Console.WriteLine($"{employee}: {latest.Date.DayMonth()} ({latest.Contract.Name})");
                }
            }
            // Possible improvement: per project
        }

        public static void ShowWrongEstimates(IEnumerable<TicketData> tickets)
        {
            Console.WriteLine("----------------");
            Console.WriteLine("Wrong estimates:");

            foreach (var ticket in tickets.Where(t => t.StoryPoints.HasValue))
            {
                if (ticket.Percent > 1.1f)
                {
                    Console.WriteLine($"{ticket.Key} actual: {ticket.Hours}h, plan: {ticket.StoryPoints * 8 * ticket.Project.DaysPerStoryPoint}h. {ticket.IssueType}: {ticket.Summary}");
                }
            }
        }
        public static void ShowTicketsWithWorkWithoutEstimates(IEnumerable<TicketData> tickets)
        {
            Console.WriteLine("----------------");
            Console.WriteLine("No estimates:");

            foreach (var ticket in tickets.Where(t => !t.StoryPoints.HasValue || t.StoryPoints.Value==0))
            {
                if (ticket.Hours > 1.0f)
                {
                    Console.WriteLine($"{ticket.Key} actual: {ticket.Hours}h. {ticket.IssueType}: {ticket.Summary}");
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
                    Console.WriteLine($"{booking.Employee}: {booking.Date.DayMonth()} ({diff}d). {booking.Contract.Name} ({booking.Contract.Id}): {startDate.DayMonth()}-{endDate.DayMonth()}");
                }
            }
        }

        public static void ShowOBookingsByEmployeeBySprint(ProjectSettings settings, IEnumerable<Booking> bookings, IEnumerable<TicketData> tickets)
        {
            Console.WriteLine("----------------");
            Console.WriteLine("Overview of bookings per sprint:");
            var employees = bookings.DistinctBy(a => a.Employee).Select(b => b.Employee).ToList();
            var contracts = settings.Contracts;
            
            foreach (var contract in contracts)
            {
                Console.WriteLine($"{contract.Name} ({contract.Id}): ");
                double totalHours = 0;
                foreach (var employee in employees)
                {
                    double hours = bookings.Aggregate(0.0f, (total, next) => next.Employee==employee && next.Contract.Id == contract.Id ? total + next.Hours : total);
                    if (hours > 0)
                    {
                        Console.WriteLine($" - {employee}: {Math.Round(hours / 8.0, 1)}d");
                    }
                    totalHours += hours;
                }
                Console.WriteLine($" = Total: {Math.Round(totalHours / 8.0, 1)}d");
            }


        }
    }
}