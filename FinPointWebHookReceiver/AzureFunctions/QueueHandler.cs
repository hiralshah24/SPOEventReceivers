using System;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Host;
using System.Collections.Specialized;
using System.Linq;
using AzureFunctions;

namespace AzureFunctions
{
    /// <summary>
    /// Handler for Azure queue containing SharePoint webhook transactions.  Receives transaction from the 
    /// queue, passes to QueueTransactionProcessor for processing.
    /// </summary>
    public static class QueueHandler
    {
        [FunctionName("QueueHandler")]
        public static void Run([QueueTrigger(Constants.WebHookStorageQueueName)]NotificationModel queueEventNotification, DateTimeOffset expirationTime
                                 , DateTimeOffset insertionTime, DateTimeOffset nextVisibleTime, string queueTrigger, string id, string popReceipt
                                 , int dequeueCount, TraceWriter log)
        {
            log.Info(System.DateTime.Now + $": Queue trigger function called for queue item {id}, notification {queueEventNotification}");

            NameValueCollection appSettings = System.Configuration.ConfigurationManager.AppSettings;

            if (!appSettings.AllKeys.Contains(Constants.TenantNameSettingName))
            {
                log.Error(System.DateTime.Now + $": Tenant Name setting {Constants.TenantNameSettingName} not found.  Queued item cannot be procssed.");
                return;
            }

            if (!appSettings.AllKeys.Contains(Constants.ClientIdSettingName))
            {
                log.Error(System.DateTime.Now + $": Client ID setting {Constants.ClientIdSettingName} not found.  Queued item cannot be procssed.");
                return;
            }

            if (!appSettings.AllKeys.Contains(Constants.ClientSecretSettingName))
            {
                log.Error(System.DateTime.Now + $": Client secret setting {Constants.ClientSecretSettingName} not found.  Queued item cannot be procssed.");
                return;
            }

            string url = String.Format("https://{0}{1}", appSettings[Constants.TenantNameSettingName], queueEventNotification.SiteUrl);
            string clientId = appSettings[Constants.ClientIdSettingName];
            string clientSecret = appSettings[Constants.ClientSecretSettingName];

            log.Info(System.DateTime.Now + $": Constructing Queue Transaction Processor for url: {url}, clientId: {clientId}, clientSecret: {clientSecret}.");

            QueueTransactionProcessor processor = new QueueTransactionProcessor(url, clientId, clientSecret, id, log);
            processor.ProcessWebHookEvents(queueEventNotification);
        }
    }
}
