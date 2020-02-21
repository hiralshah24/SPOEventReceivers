namespace AzureFunctions
{
    class Constants
    {
        /// <summary>
        /// Defines the name of the storage queue into which the webhook will place events to be processed.
        /// </summary>
        /// <remarks>
        /// Storage queue names must be all lower case.
        /// </remarks>
        public const string WebHookStorageQueueName = "webhooknotifications";

        /// <summary>
        /// The name of the web application setting in which the Azure storage account name may be found.
        /// </summary>
        public const string JobStorageSettingName = "AzureWebJobsStorage";

        /// <summary>
        /// The name of the web application setting in which the O365 tenant name may be found.
        /// </summary>
        public const string TenantNameSettingName = "TenantName";

        /// <summary>
        /// The name of the web application setting in which the O365 add-in client id may be found.
        /// </summary>
        public const string ClientIdSettingName = "ClientId";

        /// <summary>
        /// The name of the web application setting in which the O365 add-in client secret may be found.
        /// </summary>
        public const string ClientSecretSettingName = "ClientSecret";
    }
}
