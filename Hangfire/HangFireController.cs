using System;
using Hangfire;

namespace HelvertonSantos.Controllers
{
    public class HangFireController
    {
        public static void Start()
        {
            RecurringJob.AddOrUpdate<ZenviaJobsController>(x => x.SendSmsBoeExpire(), "* 7 * * 1-5", TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time"));
            RecurringJob.AddOrUpdate<ZenviaJobsController>(x => x.SendSmsBoeExpired(), "* 7 * * 1-5", TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time"));

            RecurringJob.AddOrUpdate<ECommerceJobsController>(x => x.SendOrderStatus(), "* 7 * * 1-5", TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time"));
            RecurringJob.AddOrUpdate<ECommerceJobsController>(x => x.SendStockStatus(), "* 7 * * 1-5", TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time"));
            RecurringJob.AddOrUpdate<ECommerceJobsController>(x => x.SendInvoiceStatus(), "* 7 * * 1-5", TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time"));

            RecurringJob.AddOrUpdate<PipefyJobsController>(x => x.GetPartners(), "* 7 * * 1-5", TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time"));

            RecurringJob.AddOrUpdate<EadJobsController>(x => x.GetOrders(), "* 7 * * 1-5", TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time"));

            RecurringJob.AddOrUpdate<GServiceJobsController>(x => x.UpdateDataEAD(), "* 7 * * 1-5", TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time"));

            RecurringJob.AddOrUpdate<GServiceJobsController>(x => x.UpdateDataPgrMeOperations(), "* 7 * * 1-5", TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time"));

            RecurringJob.AddOrUpdate<SalesforceJobsController>(x => x.GetSForceToken(), "* 7 * * 1-5", TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time"));
            RecurringJob.AddOrUpdate<SalesforceJobsController>(x => x.SyncAccount(), "* 7 * * 1-5", TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time"));
            RecurringJob.AddOrUpdate<SalesforceJobsController>(x => x.SyncOpportunity(), "* 7 * * 1-5", TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time"));

            RecurringJob.AddOrUpdate<DootaxJobsController>(x => x.Request(), "* 7 * * 1-5", TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time"));

            RecurringJob.AddOrUpdate<SapRpaJobsController>(x => x.PrchRtrnCompany1_To_OrdrRtrnCompany2(), "* 7 * * 1-5", TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time"));

            RecurringJob.AddOrUpdate<PepperiJobsController>(x => x.GetOrders(), "* 7 * * 1-5", TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time"));
            RecurringJob.AddOrUpdate<PepperiJobsController>(x => x.UpdateOrders(), "* 7 * * 1-5", TimeZoneInfo.FindSystemTimeZoneById("E. South America Standard Time"));
        }
    }
}
