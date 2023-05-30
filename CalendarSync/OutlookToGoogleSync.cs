using Microsoft.Office.Interop.Outlook;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Calendar.v3;
using Google.Apis.Calendar.v3.Data;
using Google.Apis.Services;
using Microsoft.EntityFrameworkCore;

namespace CalendarSync;

public class OutlookToGoogleSync
{
    private readonly MAPIFolder _outlookFolder;
    private readonly string _googleCalendarId;
    private readonly CalendarService _google;
    private readonly ApplicationDbContext _context;
    private readonly string _timezone;

    public OutlookToGoogleSync(GoogleCalendarApiConfig config)
    {
        _googleCalendarId = config.CalendarId;
        _timezone = config.Timezone;
        var credentials = GoogleCredential
            .FromFile(config.CredentialPath)
            .CreateScoped(CalendarService.Scope.Calendar);

        _google = new CalendarService(new BaseClientService.Initializer
        {
            HttpClientInitializer = credentials
        });

        var outlook = new Application();
        _outlookFolder = outlook.Session.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);

        var optionsBuilder = new DbContextOptionsBuilder<ApplicationDbContext>();
        optionsBuilder.UseSqlite("Data Source=GoogleEvents.db");
        _context = new ApplicationDbContext(optionsBuilder.Options);
    }

    /// <summary>
    /// Performs the regular sync between outlook and the local db. If there's changes to be made, it will update the local db and google calendar.
    /// </summary>
    /// <returns></returns>
    public async Task Sync()
    {
        Console.WriteLine("Starting sync...");

        // Load the current events from Outlook
        Console.WriteLine("Loading events from Outlook...");
        var outlookAppointments = GetOutlookAppointments();
        Console.WriteLine($"Loaded {outlookAppointments.Count} events from Outlook.");

        // Get the future events from the database
        Console.WriteLine("Loading events from database...");
        var dbEvents = await _context.Events.Where(e => e.IsRecurring || e.Start >= DateTime.Now.Date).ToListAsync();
        Console.WriteLine($"Loaded {dbEvents.Count} events from database.");

        // Create a dictionary of Google events by their iCalUId
        Console.WriteLine("Creating dictionary of Google events...");
        var googleEventsByIcalUId = dbEvents.ToDictionary(e => e.CalUid);

        // Iterate through the Outlook appointments
        Console.WriteLine("Iterating through Outlook appointments...");
        foreach (AppointmentItem outlookAppointment in outlookAppointments)
        {
            try
            {
                // Check if the appointment exists in the Google calendar
                var eventKey = outlookAppointment.GlobalAppointmentID;
                if (googleEventsByIcalUId.TryGetValue(eventKey, out var googleEvent))
                {
                    // The appointment exists in both Outlook and Google, check if it needs to be updated
                    if (googleEvent.Summary != outlookAppointment.Subject ||
                        googleEvent.Start != outlookAppointment.Start ||
                        googleEvent.End != outlookAppointment.End ||
                        googleEvent.Description != outlookAppointment.Body)
                    {
                        // The appointment has been updated in Outlook, update it in the Google calendar
                        Console.WriteLine($"Updating event {outlookAppointment.Subject}...");
                        await UpdateEvent(outlookAppointment, googleEvent);
                        Console.WriteLine($"Event {outlookAppointment.Subject} updated.");
                    }

                    // Remove the event from the dictionary so we can keep track of which events are deleted
                    googleEventsByIcalUId.Remove(eventKey);
                }
                else
                {
                    // The appointment doesn't exist in the Google calendar, create it
                    Console.WriteLine($"Creating event {outlookAppointment.Subject}...");
                    await CreateEvent(outlookAppointment);
                    Console.WriteLine($"Event {outlookAppointment.Subject} created.");
                }
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error creating/updating event: {ex.Message}");
            }
        }

        // Any events left in the dictionary have been deleted from Outlook, delete them from the database and the Google Calendar API
        Console.WriteLine("Checking for deleted events...");
        foreach (var googleEvent in googleEventsByIcalUId.Values)
        {
            try
            {
                Console.WriteLine($"Deleting event {googleEvent.Summary}...");
                await DeleteEvent(googleEvent);
                Console.WriteLine($"Event {googleEvent.Summary} deleted.");
            }
            catch (System.Exception ex)
            {
                Console.WriteLine($"Error deleting event: {ex.Message}");
            }
        }

        Console.WriteLine("Sync complete.");
    }

    public async Task CreateEvent(AppointmentItem outlookAppointment)
    {
        // Create the event in the Google Calendar API
        var googleCalendarEvent = ConvertToGoogleEvent(outlookAppointment);
        var createdEvent = await _google.Events.Insert(googleCalendarEvent, _googleCalendarId).ExecuteAsync();

        // Create the event in the database
        var googleEvent = new GoogleEvent
        {
            CalUid = outlookAppointment.GlobalAppointmentID,
            GoogleEventId = createdEvent.Id,
            Summary = outlookAppointment.Subject,
            Start = outlookAppointment.Start,
            End = outlookAppointment.End,
            Description = outlookAppointment.Body,
            IsRecurring = outlookAppointment.IsRecurring
        };

        _context.Events.Add(googleEvent);
        await _context.SaveChangesAsync();

        // Make sure the entry is in the db
        var entry = await _context.Events.SingleAsync(x => x.Id == googleEvent.Id);
    }

    public async Task UpdateEvent(AppointmentItem outlookAppointment, GoogleEvent googleEvent)
    {
        // Update the event in the Google Calendar API
        var googleCalendarEvent = ConvertToGoogleEvent(outlookAppointment);
        await _google.Events.Update(googleCalendarEvent, _googleCalendarId, googleEvent.GoogleEventId).ExecuteAsync();

        // Update the event in the database
        googleEvent.Summary = outlookAppointment.Subject;
        googleEvent.Start = outlookAppointment.Start;
        googleEvent.End = outlookAppointment.End;
        googleEvent.Description = outlookAppointment.Body;
        await _context.SaveChangesAsync();
    }

    public async Task DeleteEvent(GoogleEvent googleEvent)
    {
        // Delete the event from the database
        _context.Events.Remove(googleEvent);
        await _context.SaveChangesAsync();

        // Delete the event from the Google Calendar API
        await _google.Events.Delete(_googleCalendarId, googleEvent.GoogleEventId).ExecuteAsync();
    }

    private Items GetOutlookAppointments()
    {
        var filter = $"[Start] >= '{DateTime.Now.Date:g}' AND [End] <= '{DateTime.Now.AddYears(1):g}'";
        var calItems = _outlookFolder.Items;
        // calItems.IncludeRecurrences = true;
        calItems.Sort("[Start]", Type.Missing);
        return calItems.Restrict(filter);
    }

    private Event ConvertToGoogleEvent(AppointmentItem appointment)
    {
        var googleEvent = new Event
        {
            Summary = appointment.Subject,
            Start = new EventDateTime
            {
                DateTime = appointment.Start,
                TimeZone = _timezone
            },
            End = new EventDateTime
            {
                DateTime = appointment.End,
                TimeZone = _timezone
            },
            Description = appointment.Body
        };

        // Check if the appointment is recurring
        if (appointment.IsRecurring)
        {
            var recPattern = appointment.GetRecurrencePattern();
            var iCalPattern = RecurrencePatternToiCal(recPattern);
            googleEvent.Recurrence = new [] { iCalPattern };
        }

        // Map Outlook busy status to Google event status
        switch (appointment.BusyStatus)
        {
            case OlBusyStatus.olBusy:
            case OlBusyStatus.olOutOfOffice:
                googleEvent.Status = "confirmed";
                break;
            case OlBusyStatus.olTentative:
                googleEvent.Status = "tentative";
                break;
            case OlBusyStatus.olFree:
                googleEvent.Status = "transparent";
                break;
            case OlBusyStatus.olWorkingElsewhere:
                break;
        }

        // Set the event's transparency based on the busy status
        googleEvent.Transparency = appointment.BusyStatus == OlBusyStatus.olFree ? "transparent" : "opaque";

        return googleEvent;
    }

    private static string RecurrencePatternToiCal(RecurrencePattern recPattern)
    {
        var iCalPattern = "RRULE:FREQ=";

        switch (recPattern.RecurrenceType)
        {
            case OlRecurrenceType.olRecursDaily:
                iCalPattern += "DAILY";
                break;
            case OlRecurrenceType.olRecursWeekly:
                iCalPattern += "WEEKLY";
                break;
            case OlRecurrenceType.olRecursMonthly:
            case OlRecurrenceType.olRecursMonthNth:
                iCalPattern += "MONTHLY";
                break;
            case OlRecurrenceType.olRecursYearly:
            case OlRecurrenceType.olRecursYearNth:
                iCalPattern += "YEARLY";
                break;
            default:
                return "";  // Unsupported recurrence type
        }

        if (recPattern.Interval > 1)
        {
            iCalPattern += ";INTERVAL=" + recPattern.Interval;
        }

        if (recPattern.RecurrenceType != OlRecurrenceType.olRecursWeekly) return iCalPattern;
        
        var byDay = ";BYDAY=";
        if ((recPattern.DayOfWeekMask & OlDaysOfWeek.olMonday) != 0) byDay += "MO,";
        if ((recPattern.DayOfWeekMask & OlDaysOfWeek.olTuesday) != 0) byDay += "TU,";
        if ((recPattern.DayOfWeekMask & OlDaysOfWeek.olWednesday) != 0) byDay += "WE,";
        if ((recPattern.DayOfWeekMask & OlDaysOfWeek.olThursday) != 0) byDay += "TH,";
        if ((recPattern.DayOfWeekMask & OlDaysOfWeek.olFriday) != 0) byDay += "FR,";
        if ((recPattern.DayOfWeekMask & OlDaysOfWeek.olSaturday) != 0) byDay += "SA,";
        if ((recPattern.DayOfWeekMask & OlDaysOfWeek.olSunday) != 0) byDay += "SU,";

        // Trim trailing comma
        byDay = byDay.TrimEnd(',');

        iCalPattern += byDay;

        return iCalPattern;
    }
}