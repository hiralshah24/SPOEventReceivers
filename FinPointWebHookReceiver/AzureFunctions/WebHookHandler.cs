using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Azure.WebJobs.Host;
using Microsoft.WindowsAzure.Storage;
using Microsoft.WindowsAzure.Storage.Queue;
using Newtonsoft.Json;
using System;
using System.Collections.Specialized;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;

namespace AzureFunctions
{
    /// <summary>
    /// Handler for SharePoint webhooks.   Called by SharePoint.  Places received transaction into Azure queue for processing.  Returns
    /// response code back to SharePoint.
    /// </summary>
    /// <remarks>
    /// Class will renew the webhook subscription if it is close to expiration.
    /// </remarks>
    public static class WebHookHandler
    {
        [FunctionName("WebHookHandler")]
        public static async Task<HttpResponseMessage> Run([HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)]HttpRequestMessage req, TraceWriter log)
        {
            HttpResponseMessage returnResponse = new HttpResponseMessage(HttpStatusCode.OK);

            log.Info(System.DateTime.Now + ":  Starting...");
            log.Info(System.DateTime.Now + $":  Request received at {req.RequestUri}");

            // Grab the validationToken URL parameter
            string validationToken = req.GetQueryNameValuePairs()
                .FirstOrDefault(q => string.Compare(q.Key, "validationtoken", true) == 0)
                .Value;

            // If a validation token is present, we need to respond within 5 seconds by  
            // returning the given validation token. This only happens when a new web hook is being added
            if (validationToken != null)
            {
                log.Info(System.DateTime.Now + $": Validation token {validationToken} received.  Returning token to caller to complete registration...");
                returnResponse.Content = new StringContent(validationToken);
                return returnResponse;
            }

            log.Info(System.DateTime.Now + ": No validation token received.  Assuming webhook event to be handled.");

            String content = await req.Content.ReadAsStringAsync();
            if (String.IsNullOrEmpty(content))
            {
                log.Info(System.DateTime.Now + ": Content is empty.  No processing will be done.");
            }
            else
            {
                log.Info(System.DateTime.Now + $": Received payload: {content}");
                var webHookNotifications = JsonConvert.DeserializeObject<ResponseModel<NotificationModel>>(content).Value;
                log.Info(System.DateTime.Now + $": Found {webHookNotifications.Count} notifications");

                if (webHookNotifications.Count > 0)
                {
                    NameValueCollection appSettings = System.Configuration.ConfigurationManager.AppSettings;

                    if (appSettings.AllKeys.Contains(Constants.JobStorageSettingName))
                    {
                        CloudStorageAccount storageAccount = CloudStorageAccount.Parse(appSettings[Constants.JobStorageSettingName]);
                        // Get queue... create if does not exist.
                        CloudQueueClient queueClient = storageAccount.CreateCloudQueueClient();
                        CloudQueue queue = queueClient.GetQueueReference(Constants.WebHookStorageQueueName);
                        queue.CreateIfNotExists();

                        log.Info(System.DateTime.Now + ": Processing notifications...");

                        foreach (NotificationModel webHookNotification in webHookNotifications)
                        {
                            // Add message to the queue for processing.
                            string messageAsJson = JsonConvert.SerializeObject(webHookNotification);
                            log.Info(System.DateTime.Now + $": Adding message to queue: {messageAsJson}");
                            queue.AddMessage(new CloudQueueMessage(messageAsJson));
                            log.Info(System.DateTime.Now + ":  Message added.");

                            #region Rewew Webhook Subscription
                            //  If we are within 5 days of expiration, we will renew the subscription.
                            if (webHookNotification.ExpirationDateTime.AddDays(-5) < DateTime.Now)
                            {
                                log.Info(System.DateTime.Now + $": Subscription due to expire at {webHookNotification.ExpirationDateTime.ToString()}.  Renewing...");

                                if (!appSettings.AllKeys.Contains(Constants.TenantNameSettingName))
                                {
                                    log.Error(System.DateTime.Now + $": Tenant Name setting {Constants.TenantNameSettingName} not found.  Subscription cannot be renewed.");
                                    HttpResponseMessage returnMessage = new HttpResponseMessage(HttpStatusCode.InternalServerError);
                                    returnMessage.ReasonPhrase = $": Tenant Name setting {Constants.TenantNameSettingName} not found.  Subscription cannot be renewed.";
                                    return returnMessage;
                                }

                                if (!appSettings.AllKeys.Contains(Constants.ClientIdSettingName))
                                {
                                    log.Error(System.DateTime.Now + $": Client ID setting {Constants.ClientIdSettingName} not found.  Subscription cannot be renewed.");
                                    HttpResponseMessage returnMessage = new HttpResponseMessage(HttpStatusCode.InternalServerError);
                                    returnMessage.ReasonPhrase = $": Client ID setting {Constants.ClientIdSettingName} not found.  Subscription cannot be renewed.";
                                    return returnMessage;
                                }

                                if (!appSettings.AllKeys.Contains(Constants.ClientSecretSettingName))
                                {
                                    log.Error(System.DateTime.Now + $": Client secret setting {Constants.ClientSecretSettingName} not found.  Subscription cannot be renewed.");
                                    HttpResponseMessage returnMessage = new HttpResponseMessage(HttpStatusCode.InternalServerError);
                                    returnMessage.ReasonPhrase = $": Client secret setting {Constants.ClientSecretSettingName} not found.  Subscription cannot be renewed.";
                                    return returnMessage;
                                }

                                string url = String.Format("https://{0}{1}", appSettings[Constants.TenantNameSettingName], webHookNotification.SiteUrl);
                                WebHookManager manager = new WebHookManager();
                                log.Info(System.DateTime.Now + $"Calling RenewListWebHookSubscription: url={url}, Resource={webHookNotification.Resource}, RequestUri={req.RequestUri.ToString()}, subscriptionId={webHookNotification.SubscriptionId}, clientId={appSettings[Constants.ClientIdSettingName]}, clientSecret={appSettings[Constants.ClientSecretSettingName]}.");
                                Boolean renewResults = await manager.RenewListWebHookSubscription(url, webHookNotification.Resource, req.RequestUri.ToString(), webHookNotification.SubscriptionId, appSettings[Constants.ClientIdSettingName], appSettings[Constants.ClientSecretSettingName], webHookNotification.ExpirationDateTime.AddMonths(1), log);
                                if (renewResults)
                                {
                                    log.Info(System.DateTime.Now + $": Subscription renewed.");
                                }
                                else
                                {
                                    log.Error(System.DateTime.Now + $": Subscription renewal failed.");
                                }
                            }
                            #endregion
                        }
                    }
                    else
                    {
                        log.Error(System.DateTime.Now + $": Storage account setting {Constants.JobStorageSettingName} not found.  Item cannot be queued for processing.");
                        returnResponse.StatusCode = HttpStatusCode.InternalServerError;
                    }
                }
                else
                {
                    log.Info(System.DateTime.Now + ": No notifications will be queued.");
                }
            }

            // if we get here we assume the request was well received
            return returnResponse;
        }
    }
}
