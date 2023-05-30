namespace CalendarSync;

public class GoogleEvent
{
    public Guid Id { get; set; } = Guid.NewGuid();
    public string CalUid { get; set; } = string.Empty;
    public string GoogleEventId { get; set; } = string.Empty;
    public string Summary { get; set; } = string.Empty;
    public DateTime Start { get; set; }
    public DateTime End { get; set; }
    public string Description { get; set; } = string.Empty;
    public bool IsRecurring { get; set; } = false;
}