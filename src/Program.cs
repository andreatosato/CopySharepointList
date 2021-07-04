using CopySharepointList.Configurations;
using CopySharepointList.Services;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Options;
using Microsoft.Graph;
using Microsoft.Identity.Client;
using System.Net.Http.Headers;
using System.Threading.Tasks;

namespace CopySharepointList
{
    public class Program
    {
        public static void Main()
        {
            var host = new HostBuilder()
                .ConfigureFunctionsWorkerDefaults()
                .ConfigureServices(services =>
                {
                    services.AddOptions<ClientAuthenticationOptions>()
                    .Configure<IConfiguration>((settings, configuration) =>
                    {
                        configuration.GetSection("AuthConfig").Bind(settings);
                    });

                    services.AddOptions<ListConfigurations>()
                   .Configure<IConfiguration>((settings, configuration) =>
                   {
                       configuration.GetSection("ListConfig").Bind(settings);
                   });

                    services.AddSingleton<IConfidentialClientApplication>(sp =>
                    {
                        var options = sp.GetRequiredService<IOptions<ClientAuthenticationOptions>>()?.Value;
                        return ConfidentialClientApplicationBuilder
                        .Create(options.ClientId)
                        .WithClientSecret(options.ClientSecret)
                        .WithTenantId(options.TenantId)
                        .Build();
                    });

                    services.AddScoped<GraphServiceClient>((sp) =>
                    {
                        var msalClient = sp.GetRequiredService<IConfidentialClientApplication>();
                        var authResult = msalClient
                            .AcquireTokenForClient(new[] { "https://graph.microsoft.com/.default" })
                            .ExecuteAsync()
                            .GetAwaiter().GetResult();

                        return new GraphServiceClient(new DelegateAuthenticationProvider(
                          (requestMessage) =>
                              {
                                  requestMessage.Headers.Authorization = new AuthenticationHeaderValue("bearer", authResult.AccessToken);
                                  return Task.CompletedTask;
                              }), null);
                    });

                    services.AddScoped<IReaderFields, ReaderFields>();
                    services.AddScoped<ISiteService, SiteService>();
                })
                .Build();

            host.Run();
        }
    }
}