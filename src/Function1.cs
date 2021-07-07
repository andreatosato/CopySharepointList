//using Microsoft.Azure.Functions.Worker;
//using Microsoft.Extensions.Logging;
//using System;

//namespace CopySharepointList
//{
//    public static class Function1
//    {
//        [Function("Function1")]
//        public static void Run([TimerTrigger("%Cron%")] MyInfo myTimer, FunctionContext context)
//        {
//            var logger = context.GetLogger("Function1");
//            logger.LogInformation($"C# Timer trigger function executed at: {DateTime.Now}");
//        }
//    }

//    public class MyInfo
//    {
//        public MyScheduleStatus ScheduleStatus { get; set; }

//        public bool IsPastDue { get; set; }
//    }

//    public class MyScheduleStatus
//    {
//        public DateTime Last { get; set; }

//        public DateTime Next { get; set; }

//        public DateTime LastUpdated { get; set; }
//    }
//}
