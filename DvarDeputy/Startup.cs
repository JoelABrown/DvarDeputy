using Microsoft.AspNetCore.Builder;
using Microsoft.Extensions.DependencyInjection;

namespace Mooseware.DvarDeputy;

internal class Startup
{
    public static void ConfigureServices(IServiceCollection services)
    {
        services.AddControllers();
    }

    public static void Configure(IApplicationBuilder app)
    {
        app.UseRouting();

        app.UseEndpoints(routes =>
        {
            routes.MapControllers();
        });
    }
}
