using Microsoft.EntityFrameworkCore;

namespace CalendarSync;

public sealed class ApplicationDbContext : DbContext
{
    public DbSet<GoogleEvent> Events { get; set; }

    public ApplicationDbContext(DbContextOptions<ApplicationDbContext> options) : base(options)
    {
        Database.EnsureCreated();
    }
}
