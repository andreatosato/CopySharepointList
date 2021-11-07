using CopySharepointList.Services;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;
using Microsoft.Extensions.Logging;
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
            var logger = executionContext.GetLogger("ScheduleCopyManual");
            logger.LogInformation("Start execution ScheduleCopyManual");

            await siteService.ExecuteAsync();

            logger.LogInformation("End execution ScheduleCopyManual");
            return new OkResult();
        }

        //[Function("ScheduleCopyTimer")]
        public async Task RunTimer([TimerTrigger("%Cron%")] object myTimer, FunctionContext context)
        {
            var logger = context.GetLogger("ScheduleCopyTimer");
            logger.LogInformation("Start execution ScheduleCopyTimer");

            await siteService.ExecuteAsync();

            logger.LogInformation("End execution ScheduleCopyTimer");
        }
    }
}
