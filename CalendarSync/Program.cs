using CalendarSync;
using Microsoft.Extensions.Configuration;

Console.WriteLine("Initializing Settings...");

var configuration = new ConfigurationBuilder()
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
    .AddJsonFile($"appsettings.Development.json", optional: true)
    .Build();

var googleConfig = new GoogleCalendarApiConfig();
configuration.GetSection("GoogleCalendarApi").Bind(googleConfig);

Console.WriteLine("Initializing Outlook/Google Sync Client...");
var outlook = new OutlookToGoogleSync(googleConfig);
Console.WriteLine("Initialized Outlook/Google Sync Client...");

while (true)
{
    Console.WriteLine("Running sync...");
    await outlook.Sync();
    Console.WriteLine("Waiting 60 seconds to continue...");
    await Task.Delay(TimeSpan.FromSeconds(60)); // wait for a minute before the next iteration
}
