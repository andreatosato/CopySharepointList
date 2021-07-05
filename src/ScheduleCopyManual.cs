using CopySharepointList.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using System;
using System.Threading.Tasks;

namespace CopySharepointList
{
    public class ScheduleCopyManual
    {
        private readonly ISiteService siteService;

        public ScheduleCopyManual(ISiteService siteService)
        {
            this.siteService = siteService ?? throw new ArgumentNullException(nameof(siteService));
        }

        [Function("ScheduleCopyManual")]
        public async Task<IActionResult> Run([HttpTrigger(AuthorizationLevel.Function, "get", "post")] HttpRequestData req,
            FunctionContext executionContext)
        {
            await siteService.ExecuteAsync();
            return new OkResult();
        }
    }
}
