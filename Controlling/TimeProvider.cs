public class TimeProvider: ITimeProvider
{
    public static ITimeProvider Instance { get; set; } = new TimeProvider();
    public DateTime Now => DateTime.Now;
}

public interface ITimeProvider
{
    DateTime Now { get; }
}