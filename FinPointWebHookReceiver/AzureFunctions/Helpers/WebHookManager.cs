using Microsoft.Azure.WebJobs.Host;
using Newtonsoft.Json;
using SharePointPnP.IdentityModel.Extensions.S2S.Protocols.OAuth2;
using System;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;

namespace AzureFunctions
{
    /// <summary>
    /// Class with methods to manage SharePoint web hooks
    /// </summary>
    public class WebHookManager
    {
        public async Task<bool> RenewListWebHookSubscription(string siteUrl, string listId, String subscriptionUrl, String subscriptionId, String clientId, String clientSecret, DateTime newExpirationDate, TraceWriter log)
        {
            log.Info(System.DateTime.Now + $"RenewListWebHookSubscription: siteUrl={siteUrl}, listId={listId}, subscriptionUrl={subscriptionUrl}, subscriptionId={subscriptionId}, clientId={clientId}, clientSecret={clientSecret}.");

            const string SHAREPOINT_PRINCIPAL = "00000003-0000-0ff1-ce00-000000000000";

            Uri siteUri = new Uri(siteUrl);
            log.Info(System.DateTime.Now + $"RenewListWebHookSubscription: siteUri={siteUri}");

            String siteUriAuthority = new Uri(siteUrl).Authority;
            log.Info(System.DateTime.Now + $"RenewListWebHookSubscription: siteUriAuthority={siteUriAuthority}");

            String realm = TokenHelper.GetRealmFromTargetUrl(new Uri(siteUrl));
            log.Info(System.DateTime.Now + $"RenewListWebHookSubscription: realm={realm}");

            OAuth2AccessTokenResponse accessToken = TokenHelper.GetAppOnlyAccessToken(SHAREPOINT_PRINCIPAL, siteUriAuthority, realm);
            log.Info(System.DateTime.Now + $"RenewListWebHookSubscription: accessToken={accessToken}");

            Task<bool> updateResult = Task.WhenAny(
                this.UpdateListWebHookAsync(
                    siteUrl,
                    listId,
                    subscriptionId,
                    subscriptionUrl,
                    newExpirationDate,
                    accessToken.AccessToken, log)
                ).Result;

            if (updateResult.Result == false)
            {
                throw new Exception(String.Format("The expiration date of web hook {0} with endpoint {1} could not be updated", subscriptionId, subscriptionUrl));
            }
            else
            {
                return await updateResult;
            }
        }

        #region Update a list web hook
        /// <summary>
        /// Updates the expiration datetime (and notification URL) of an existing SharePoint list web hook
        /// </summary>
        /// <param name="siteUrl">Url of the site holding the list</param>
        /// <param name="listId">Id of the list</param>
        /// <param name="subscriptionId">Id of the web hook subscription that we need to update</param>
        /// <param name="webHookEndPoint">Url of the web hook service endpoint (the one that will be called during an event)</param>
        /// <param name="expirationDateTime">New web hook expiration date</param>
        /// <param name="accessToken">Access token to authenticate against SharePoint</param>
        /// <returns>true if succesful, exception in case something went wrong</returns>
        private async Task<bool> UpdateListWebHookAsync(string siteUrl, string listId, string subscriptionId, string webHookEndPoint, DateTime expirationDateTime, string accessToken, TraceWriter log)
        {
            log.Info(System.DateTime.Now + $"UpdateListWebHookAsync: siteUrl={siteUrl}, listId={listId}, subscriptionId={subscriptionId}, webHookEndPoint={webHookEndPoint}, expirationDateTime={expirationDateTime}, accessToken={accessToken}");
            using (var httpClient = new HttpClient())
            {
                string requestUrl = String.Format("{0}/_api/web/lists('{1}')/subscriptions('{2}')", siteUrl, listId, subscriptionId);
                log.Info(System.DateTime.Now + $"UpdateListWebHookAsync: requestUrl={requestUrl}");
                HttpRequestMessage request = new HttpRequestMessage(new HttpMethod("PATCH"), requestUrl);
                request.Headers.Accept.Add(new MediaTypeWithQualityHeaderValue("application/json"));
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

                request.Content = new StringContent(JsonConvert.SerializeObject(
                    new SubscriptionModel()
                    {
                        NotificationUrl = webHookEndPoint,
                        ExpirationDateTime = expirationDateTime.ToUniversalTime(),
                    }),
                    Encoding.UTF8, "application/json");

                HttpResponseMessage response = await httpClient.SendAsync(request);

                if (response.StatusCode != System.Net.HttpStatusCode.NoContent)
                {
                    String responseContent = await response.Content.ReadAsStringAsync();
                    log.Error(System.DateTime.Now + $"RenewListWebHookSubscription: response.StatusCode={response.StatusCode}, content={responseContent}");
                    // oops...something went wrong, maybe the web hook does not exist?
                    throw new Exception(await response.Content.ReadAsStringAsync());
                }
                else
                {
                    return await Task.Run(() => true);
                }
            }
        }
        #endregion
    }
}
