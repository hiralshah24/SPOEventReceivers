using Microsoft.Azure.WebJobs.Host;
using System;
using Microsoft.SharePoint.Client;
using System.Collections.Generic;
using System.Linq;
using AzureFunctions;
using OfficeDevPnP.Core;

namespace AzureFunctions
{
    public class QueueTransactionProcessor
    {
        const String ProcessName = "Queue Transaction Processor";

        TraceWriter log = null;
        ClientContext sharePointContext = null;
        private WebhookHistoryList historyList;
        private WebHookList webHookList;
        private ExpenseSummaryList expenseSummaryList;
        Boolean initialized = false;

        /// <summary>
        /// Constructs a new instance of <see cref="QueueTransactionProcessor"/> using application client id and secret.
        /// </summary>
        /// <param name="siteUrl"></param>
        /// <param name="clientId"></param>
        /// <param name="clientSecret"></param>
        /// <param name="id">A <see cref="String"/> containing a value that will uniquely identify this <see cref="NotificationProcessor"/>
        /// <param name="log">A <see cref="TraceWriter"/> providing logging support.</param>
        public QueueTransactionProcessor(String siteUrl, String clientId, String clientSecret, String id, TraceWriter log)
        {
            this.log = log;
            this.log.Info(System.DateTime.Now + $": Constructing.  SiteUrl: {siteUrl}, clientId: {clientId}, clientSecret: {clientSecret}.");
            AuthenticationManager am = new AuthenticationManager();
            this.sharePointContext = am.GetAppOnlyAuthenticatedContext(siteUrl, clientId, clientSecret);
            this.log.Info(System.DateTime.Now + $": SharePoint context obtained.  Initializing...");
            Boolean loadListResults = LoadLists();
            if (loadListResults)
            {
                this.log.Info("Initialization complete");
                this.initialized = true;
            }
        }

        private Boolean LoadLists()
        {
            this.webHookList = new WebHookList(this.sharePointContext);
            this.historyList = new WebhookHistoryList(this.sharePointContext);

            if (this.webHookList == null)
            {
                this.log.Error(String.Format("Failed to load Web Hook list. Terminating."));
                return false;
            }
            else if (this.historyList == null)
            {
                this.log.Error(String.Format("Failed to load Web Hook history list. Terminating."));
                return false;
            }

            return true;
        }

        public void ProcessWebHookEvents(NotificationModel eventNotification)
        {
            if (!this.initialized)
            {
                this.log.Error($"Synchronizer not property initialized. See related messages.");
                return;
            }
            this.log.Info($"Processing queue event notification: {eventNotification}");
            DateTime startTime = DateTime.Now;
            ListCollection lists = sharePointContext.Web.Lists;
            Guid listId = new Guid(eventNotification.Resource);
            IEnumerable<List> eventListQueryResults = sharePointContext.LoadQuery<List>(lists.Where(lst => lst.Id == listId));
            try
            {
                sharePointContext.ExecuteQueryRetry();
            }
            catch (Exception ex)
            {
                this.log.Error($"Exception:{ex.Message}");
            }
            List changeList = eventListQueryResults.FirstOrDefault();

            if (changeList == null)
            {
                this.log.Warning($"List Id {eventNotification.Resource} does not exist on site {sharePointContext.Web.Url}. Likely has been deleted since event was queued.");
                return;
            }

            //Could easily expand this to handle other lists.
            if (String.Compare(changeList.Title, WebHookList.DefaultListTitle, StringComparison.InvariantCultureIgnoreCase) != 0)
            {
                this.log.Error($"Title of event list {changeList.Title} does not match expected list title {WebHookList.DefaultListTitle}. This handler has likely been attached to the incorrect list.  Received notification cannot be processed.");
                return;
            }

            this.log.Info($"Retrieving history token");
            this.historyList = new WebhookHistoryList(this.sharePointContext);
            String lastChangeTokenValue = this.historyList.GetLastToken(ProcessName);
            ChangeToken lastChangeToken = new ChangeToken();

            if (String.IsNullOrEmpty(lastChangeTokenValue))
            {
                // See https://blogs.technet.microsoft.com/stefan_gossner/2009/12/04/content-deployment-the-complete-guide-part-7-change-token-basics/
                // The format of the string is semicolon delimited with the following pieces of information in order
                // Version number 
                // A number indicating the change scope: 0 – Content Database, 1 – site collection, 2 – site, 3 – list. 
                // GUID representing the scope ID of the change token (e.g., the list ID)
                // Time (in UTC) when the change occurred
                // Number of the change relative to other changes
                // If there is no token, we will default to the last 10 minutes under the assumption that that is the time that has elapsed 
                //   since a change was made and the trigger got called.
                lastChangeToken.StringValue = string.Format("1;3;{0};{1};-1", eventNotification.Resource, DateTime.Now.AddMinutes(-2).ToUniversalTime().Ticks.ToString());
                this.log.Info($"No history token found.  Initializing to {lastChangeToken.StringValue}");
            }
            else
            {
                lastChangeToken.StringValue = lastChangeTokenValue;
                this.log.Info($"History token retrieved : {lastChangeToken.StringValue}");
            }

            ChangeQuery changeQuery = new ChangeQuery(false, false);
            changeQuery.Item = true;
            changeQuery.Add = true;
            changeQuery.Update = true;
            changeQuery.FetchLimit = 2000; // Max value is 2000, default = 1000
            //  Create a new change token = last change token + 1 millisecond. We don't want to re-process the last change.
            lastChangeToken.StringValue = $"{lastChangeToken.GetVersion()};{lastChangeToken.GetChangeScope()};{lastChangeToken.GetScopeId()};{lastChangeToken.GetDate().AddMilliseconds(1).Ticks.ToString()};{lastChangeToken.GetChangeNumber()}";
            changeQuery.ChangeTokenStart = lastChangeToken;
            this.log.Info($"Submitted token {lastChangeToken.StringValue} for change query.");
            ChangeCollection listChanges = changeList.GetChanges(changeQuery);
            this.sharePointContext.Load(listChanges);
            this.sharePointContext.ExecuteQueryRetry();

            HashSet<Int32> itemIdstoBeProcessed = new HashSet<Int32>();
            Int32 itemsRetrieved = 0;
            Int32 itemsProcessed = 0;
            this.log.Info($"ListChanges {listChanges.Count()} ");
            //  For each change returned by our change query...
            foreach (Change currentListItemChange in listChanges)
            {
                itemsRetrieved++;
                if (currentListItemChange.GetType() == typeof(Microsoft.SharePoint.Client.ChangeItem))
                {
                    ChangeItem changedItem = currentListItemChange.TypedObject as ChangeItem;
                    lastChangeToken = changedItem.ChangeToken;
                    //if (changedItem.ChangeType == ChangeType.Add || changedItem.ChangeType == ChangeType.Update)
                    {
                        if (itemIdstoBeProcessed.Contains(changedItem.ItemId))
                        {
                            this.log.Info($"Duplicate {changedItem.ChangeType} for item ID {changedItem.ItemId} received from SharePoint.");
                        }
                        else
                        {
                            itemIdstoBeProcessed.Add(changedItem.ItemId);
                            itemsProcessed++;
                            this.log.Info($"Item change received.  Change Type: { currentListItemChange.ChangeType}, Change token: { currentListItemChange.ChangeToken.StringValue}, Item Id: {changedItem.ItemId}");
                        }
                    }
                    /*else
                    {
                        this.log.Info($"Transction { currentListItemChange.ChangeType} on Person Event ID {changedItem.ItemId} ignored.");
                    }*/
                }
                else
                {
                    this.log.Warning($"Alert change received.  Will NOT be processed.  Change Type: { currentListItemChange.ChangeType}, Change token: { currentListItemChange.ChangeToken.StringValue}");
                }
            }

            ProcessNotificationsByItemId(itemIdstoBeProcessed, this.webHookList, this.log);

            Boolean historyAddResults = this.historyList.TryAddItem(ProcessName, itemsRetrieved, itemsProcessed, startTime, DateTime.Now, lastChangeToken.StringValue);

            if (historyAddResults == false)
            {
                this.log.Warning($"Notification history addition failed.");
            }

            if (this.log != null) { this.log.Info(System.DateTime.Now + $": Complete."); }
            this.log.Info($"Complete.");
        }

        /// <summary>
        /// Given a <see cref="List"/> of Notification item ids, processes all items that have not already been processed.
        /// </summary>
        /// <param name="itemIdsToBeProcessed"></param>
        /// <param name="webHookList"></param>
        /// <param name="logger"></param>
        internal void ProcessNotificationsByItemId(HashSet<Int32> itemIdsToBeProcessed, WebHookList webHookList, TraceWriter logger)
        {
            List<ListItem> unProcessedEventItems = webHookList.GetItemsById(itemIdsToBeProcessed);

            if (itemIdsToBeProcessed.Count() == unProcessedEventItems.Count())
            {
                logger.Info($"{itemIdsToBeProcessed.Count()} Id(s) passed for event lookup.  {unProcessedEventItems.Count()} item(s) received back.");
            }
            else
            {
                logger.Warning($"{itemIdsToBeProcessed.Count()} Id(s) passed for event lookup.  {unProcessedEventItems.Count()} item(s) received back.");
            }

            int itemsProcessed = ProcessNotificationListItems(webHookList, logger, unProcessedEventItems.AsQueryable());

            /*if (itemsProcessed == 0)
            {
                logger.Info("No unprocessed notification found.");
            }
            else
            {
                logger.Info(itemsProcessed + " notification processed.");
            }*/
        }

        /// <summary>
        /// Processes all <see cref="ListItem"/>s (from the web hook list) that have not already been processed.
        /// </summary>
        /// <param name="webHookList"></param>
        /// <param name="logger"></param>
        /// <param name="unProcessedItems"></param>
        /// <returns></returns>
        private int ProcessNotificationListItems(WebHookList webHookList, TraceWriter logger, IQueryable<ListItem> unProcessedItems)
        {
            Int32 itemsProcessed = 0;
            Dictionary<string, string> SourceItem = new Dictionary<string, string>();
            foreach (ListItem currentItem in unProcessedItems)
            {
                logger.Info("total unprocessed items : " + unProcessedItems.Count());
                try
                {
                    SourceItem = webHookList.ProcessItem(currentItem);
                    FieldLookupValue companyLookup = currentItem[WebHookList.ChargedCompany] as FieldLookupValue;
                    FieldLookupValue costCenterLookup = currentItem[WebHookList.CostCenter] as FieldLookupValue;

                    logger.Info("Expensedata : " + currentItem[WebHookList.EventCategory]);
                    logger.Info("Expensedata : " + costCenterLookup.LookupValue);
                    logger.Info("Expensedata : " + currentItem[WebHookList.FiscalYear].ToString());
                    this.expenseSummaryList = new ExpenseSummaryList(this.sharePointContext);
                    //Boolean summaryAddResult = expenseSummaryList.TryAddItem(companyLookup.LookupId,costCenterLookup.LookupId, currentItem[WebHookList.FiscalYear].ToString());
                    Boolean summaryAddResult = expenseSummaryList.TryAddItem(currentItem);
                    itemsProcessed++;
                }
                catch (Exception ex)
                {
                    log.Error("Error:" + ex.Message);
                }

            }

            if (itemsProcessed > 0)
            {
                this.sharePointContext.ExecuteQuery();
            }

            return itemsProcessed;
        }
    }
}
