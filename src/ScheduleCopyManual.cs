using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CopySharepointList
{
    public class ScheduleCopyManual
    {
        private readonly GraphServiceClient graphServiceClient;
        private readonly string[] listsToCopy;
        public ScheduleCopyManual(GraphServiceClient graphServiceClient, IConfiguration configuration)
        {
            this.graphServiceClient = graphServiceClient;
            listsToCopy = configuration.GetValue<string>("ListsToCopy").Split(";");
        }

        [Function("ScheduleCopyManual")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req,
            FunctionContext executionContext)
        {
            var listsMaster = await graphServiceClient.Sites["f6216b26-a584-4249-b9ac-3215bb884a89"]
                .Lists
                .Request()
                .GetAsync();

            do
            {
                foreach (var l in listsMaster.CurrentPage)
                {
                    if (listsToCopy.Contains(l.Name))
                    {
                        Console.WriteLine($"{l.Name}");

                        var queryOptions = new List<QueryOption>()
                        {
                            new QueryOption("expand", "fields")
                        };

                        var listItems = await graphServiceClient
                            .Sites["f6216b26-a584-4249-b9ac-3215bb884a89"]
                            .Lists[l.Id]
                            .Items
                            .Request(queryOptions).GetAsync();
                        foreach (var row in listItems.CurrentPage)
                        {
                            var title = row.Fields.AdditionalData["Title"];
                            var ragioneSociale = row.Fields.AdditionalData["RagioneSociale"];
                            Console.WriteLine($"{title}{ragioneSociale}");

                        }
                        //var f = items.CurrentPage[0].Fields;
                    }
                }
                if (listsMaster.NextPageRequest != null)
                    await listsMaster.NextPageRequest.GetAsync();
            }
            while (listsMaster.NextPageRequest != null);

            return new OkResult();
        }
    }
}
