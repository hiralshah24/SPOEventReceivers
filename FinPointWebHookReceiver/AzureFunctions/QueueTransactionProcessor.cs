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
        public QueueTransactionProcessor(String siteUrl, String clientId, String clientSecret, String id, TraceWriter log, Guid listId)
        {
            this.log = log;
            this.log.Info(System.DateTime.Now + $": Constructing.  SiteUrl: {siteUrl}, clientId: {clientId}, clientSecret: {clientSecret}.");
            AuthenticationManager am = new AuthenticationManager();
            this.sharePointContext = am.GetAppOnlyAuthenticatedContext(siteUrl, clientId, clientSecret);
            this.log.Info(System.DateTime.Now + $": SharePoint context obtained.  Initializing...");
            Boolean loadListResults = LoadLists(listId);
            if (loadListResults)
            {
                this.log.Info("Initialization complete");
                this.initialized = true;
            }
        }

        private Boolean LoadLists(Guid listId)
        {
            this.webHookList = new WebHookList(this.sharePointContext, listId);
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

            if (!String.IsNullOrEmpty(lastChangeTokenValue) && lastChangeTokenValue.IndexOf(eventNotification.Resource) != -1)
            {
                lastChangeToken.StringValue = lastChangeTokenValue;
                this.log.Info($"History token retrieved : {lastChangeToken.StringValue}");

            }
            else
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

            ChangeQuery changeQuery = new ChangeQuery(false, false);
            changeQuery.Item = true;
            changeQuery.Add = true;
            changeQuery.Update = true;
            changeQuery.FetchLimit = 2000; // Max value is 2000, default = 1000
            //  Create a new change token = last change token + 1 millisecond. We don't want to re-process the last change.
            lastChangeToken.StringValue = $"{lastChangeToken.GetVersion()};{lastChangeToken.GetChangeScope()};{lastChangeToken.GetScopeId()};{lastChangeToken.GetDate().AddMilliseconds(1).Ticks.ToString()};{lastChangeToken.GetChangeNumber()}";
            changeQuery.ChangeTokenStart = lastChangeToken;
            this.log.Info($"Submitted token {lastChangeToken.StringValue} for change query.");
            this.log.Info($"ChangeList:{changeList.Title} and Defaulttitle {WebHookList.DefaultListTitle}");
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
            List<ListItem> returnItems = new List<ListItem>();
            try
            {
                List<ListItem> unProcessedEventItems = webHookList.GetItemsById(itemIdsToBeProcessed);
                //List list = this.sharePointContext.Web.Lists.GetByTitle(webHookList.ListTitle);
                //this.sharePointContext.Load(list);
                //this.sharePointContext.ExecuteQuery();
                //log.Info("Listtitle : " + list.Title);               

                //foreach (Int32 currentItemId in itemIdsToBeProcessed)
                //{

                //    ListItem eventItem = list.GetItemById(currentItemId);
                //    log.Info("itemId : " + currentItemId);
                //    this.sharePointContext.Load(eventItem, item => item.Id, item => item["Title"], item => item["Created"], item => item["Author"], item => item["Modified"], item => item["ContentTypeId"]
                //                                                                , item => item[WebHookList.WebhookProcessedDate], item => item[WebHookList.CostCenter], item => item[WebHookList.ChargedCompany], item => item[WebHookList.ExpenseCategory]
                //                                                                , item => item[WebHookList.GLAccount], item => item[WebHookList.JanAmt], item => item[WebHookList.FebAmt], item => item[WebHookList.MarAmt]
                //                                                                , item => item[WebHookList.AprAmt], item => item[WebHookList.MayAmt], item => item[WebHookList.JunAmt], item => item[WebHookList.JunAmt]
                //                                                                , item => item[WebHookList.AugAmt], item => item[WebHookList.SepAmt], item => item[WebHookList.OctAmt], item => item[WebHookList.NovAmt]
                //                                                                , item => item[WebHookList.DecAmt], item => item[WebHookList.FYNext1], item => item[WebHookList.FYNext2], item => item[WebHookList.FiscalYear]);//, item => item[WebHookList.CompanyCode]);//, item => item[CostCenterCode]);

                //    this.sharePointContext.ExecuteQuery();
                //    if (eventItem[WebHookList.ExpenseCategory].ToString() == "IT")
                //    {
                //        returnItems.Add(eventItem);
                //    }
                //}
                //List<ListItem> unProcessedEventItems = returnItems;

                if (itemIdsToBeProcessed.Count() == unProcessedEventItems.Count())
                {
                    logger.Info($"{itemIdsToBeProcessed.Count()} Id(s) passed for event lookup.  {unProcessedEventItems.Count()} item(s) received back.");
                }
                else
                {
                    logger.Warning($"{itemIdsToBeProcessed.Count()} Id(s) passed for event lookup.  {unProcessedEventItems.Count()} item(s) received back.");
                }

                int itemsProcessed = ProcessNotificationListItems(webHookList, logger, unProcessedEventItems.AsQueryable());

                if (itemsProcessed == 0)
                {
                    logger.Info("No unprocessed notification found.");
                }
                else
                {
                    logger.Info(itemsProcessed + " notification processed.");
                }
            }
            catch (Microsoft.SharePoint.Client.ServerException ex)
            {
                log.Info("Message :" + ex.Message);
            }
            catch (Exception ex)
            {
                log.Info("Message :" + ex.Message);
            }
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
            Dictionary<string, string> SourceData = new Dictionary<string, string>();
            foreach (ListItem currentItem in unProcessedItems)
            {
                logger.Info("total unprocessed items : " + unProcessedItems.Count());
                try
                {
                    SourceData = webHookList.ProcessItem(currentItem);
                    #region ProcessExpenseItem
                    //FieldLookupValue GLAccountLookup = currentItem[WebHookList.GLAccount] as FieldLookupValue;
                    //FieldLookupValue CompanyLookup = currentItem[WebHookList.ChargedCompany] as FieldLookupValue;
                    //FieldLookupValue CostCenterLookup = currentItem[WebHookList.CostCenter] as FieldLookupValue;
                    //Dictionary<string, string> SourceData = new Dictionary<string, string>();

                    //Folder folder = this.sharePointContext.Web.GetFolderByServerRelativeUrl(this.sharePointContext.Url + "/Lists/" +
                    //                WebHookList.DefaultListTitle + "/" + currentItem[WebHookList.FiscalYear] + "/" +
                    //                CompanyLookup.LookupValue.Split('-')[0].Trim() + "-" + CostCenterLookup.LookupValue.Split('-')[0].Trim());
                    //this.log.Info("Foldern:" + this.sharePointContext.Url + "/Lists/" +
                    //                WebHookList.DefaultListTitle + "/" + currentItem[WebHookList.FiscalYear] + "/" +
                    //                CompanyLookup.LookupValue.Split('-')[0].Trim() + "-" + CostCenterLookup.LookupValue.Split('-')[0].Trim());
                    //this.sharePointContext.Load(folder);
                    //this.sharePointContext.ExecuteQuery();
                    //if (folder.Exists)
                    //{

                    //    CamlQuery query = new CamlQuery();

                    //    query.FolderServerRelativeUrl = folder.ServerRelativeUrl;
                    //    this.log.Info("FolderURL:" + query.FolderServerRelativeUrl);
                    //    query.ViewXml = "<View Scope=\"RecursiveAll\"> " +
                    //                        "<Query>" +
                    //                            "<Where>"
                    //                            + "<And>"
                    //                                + "<Eq><FieldRef Name='GL_x0020_Account'/><Value Type='Lookup'>" + GLAccountLookup.LookupValue + "</Value></Eq>"
                    //                                + "<Eq><FieldRef Name='Expense_Category' /><Value Type='Choice'>IT</Value></Eq>"
                    //                            + "</And>"
                    //                            + "</Where>"
                    //                        + "</Query>"
                    //                    + "</View>";
                    //    ListItemCollection colItems = this.sharePointContext.Web.Lists.GetByTitle(WebHookList.DefaultListTitle).GetItems(query);
                    //    this.sharePointContext.Load(colItems);
                    //    this.sharePointContext.ExecuteQuery();
                    //    Int64 jan = 0;
                    //    Int64 feb = 0;
                    //    Int64 mar = 0;
                    //    Int64 apr = 0;
                    //    Int64 may = 0;
                    //    Int64 jun = 0;
                    //    Int64 jul = 0;
                    //    Int64 aug = 0;
                    //    Int64 sep = 0;
                    //    Int64 oct = 0;
                    //    Int64 nov = 0;
                    //    Int64 dec = 0;
                    //    Int64 fynext = 0;
                    //    Int64 fynextnext = 0;
                    //    if (colItems != null && colItems.Count > 0)
                    //    {
                    //        foreach (ListItem colItem in colItems)
                    //        {
                    //            jan += Convert.ToInt64(colItem[WebHookList.JanAmt]);
                    //            feb += Convert.ToInt64(colItem[WebHookList.FebAmt]);
                    //            mar += Convert.ToInt64(colItem[WebHookList.MarAmt]);
                    //            apr += Convert.ToInt64(colItem[WebHookList.AprAmt]);
                    //            may += Convert.ToInt64(colItem[WebHookList.MayAmt]);
                    //            jun += Convert.ToInt64(colItem[WebHookList.JunAmt]);
                    //            jul += Convert.ToInt64(colItem[WebHookList.JulAmt]);
                    //            aug += Convert.ToInt64(colItem[WebHookList.AugAmt]);
                    //            sep += Convert.ToInt64(colItem[WebHookList.SepAmt]);
                    //            oct += Convert.ToInt64(colItem[WebHookList.OctAmt]);
                    //            nov += Convert.ToInt64(colItem[WebHookList.NovAmt]);
                    //            dec += Convert.ToInt64(colItem[WebHookList.DecAmt]);
                    //            fynext += Convert.ToInt64(colItem[WebHookList.FYNext1]);
                    //            fynextnext += Convert.ToInt64(colItem[WebHookList.FYNext2]);
                    //        }

                    //        SourceData.Add("Phase", webHookList.Phase);
                    //        SourceData.Add(WebHookList.ChargedCompany, CompanyLookup.LookupId.ToString());
                    //        SourceData.Add(WebHookList.CostCenter, CostCenterLookup.LookupId.ToString());
                    //        SourceData.Add(WebHookList.GLAccount, GLAccountLookup.LookupId.ToString());
                    //        SourceData.Add(WebHookList.FiscalYear, currentItem[WebHookList.FiscalYear].ToString());
                    //        SourceData.Add(WebHookList.JanAmt, jan.ToString());
                    //        SourceData.Add(WebHookList.FebAmt, feb.ToString());
                    //        SourceData.Add(WebHookList.MarAmt, mar.ToString());
                    //        SourceData.Add(WebHookList.AprAmt, apr.ToString());
                    //        SourceData.Add(WebHookList.MayAmt, may.ToString());
                    //        SourceData.Add(WebHookList.JunAmt, jun.ToString());
                    //        SourceData.Add(WebHookList.JulAmt, jul.ToString());
                    //        SourceData.Add(WebHookList.AugAmt, aug.ToString());
                    //        SourceData.Add(WebHookList.SepAmt, sep.ToString());
                    //        SourceData.Add(WebHookList.OctAmt, oct.ToString());
                    //        SourceData.Add(WebHookList.NovAmt, nov.ToString());
                    //        SourceData.Add(WebHookList.DecAmt, dec.ToString());
                    //        SourceData.Add(WebHookList.FYNext1, fynext.ToString());
                    //        SourceData.Add(WebHookList.FYNext2, fynextnext.ToString());
                    //    }
                    #endregion
                    logger.Info("Items:" + SourceData.Count());
                    this.expenseSummaryList = new ExpenseSummaryList(this.sharePointContext);
                    if (SourceData != null && SourceData.Count > 0)

                    {
                        Boolean summaryAddResult = expenseSummaryList.TryAddUpdateItem(SourceData);
                        #region SummaryFunction
                        //    FieldLookupValue CompanyLookup = new FieldLookupValue();
                        //    CompanyLookup.LookupId = int.Parse(sourceItem[WebHookList.ChargedCompany]);
                        //    FieldLookupValue CostCenterLookup = new FieldLookupValue();
                        //    CostCenterLookup.LookupId = int.Parse(sourceItem[WebHookList.CostCenter]);
                        //    FieldLookupValue GLAccountLookup = new FieldLookupValue();
                        //    GLAccountLookup.LookupId = int.Parse(sourceItem[WebHookList.GLAccount]);
                        //    CamlQuery targetQuery = new CamlQuery();

                        //    targetQuery.ViewXml = "<View>"
                        //                            + "<Query>"
                        //                            + "<Where>"
                        //                                + "<And>"
                        //                                    + "<Eq><FieldRef Name='Owning_x0020_Company' LookupId='TRUE' /><Value Type='Lookup'>" + int.Parse(sourceItem[WebHookList.ChargedCompany]) + "</Value></Eq>"
                        //                                    + "<And>"
                        //                                        + "<Eq><FieldRef Name='Fiscal_Year'/><Value Type='Number'>" + int.Parse(sourceItem[WebHookList.FiscalYear]) + "</Value></Eq>"
                        //                                            + "<And>"
                        //                                                + "<Eq><FieldRef Name='Owning_x0020_CostCenter' LookupId='TRUE'/><Value Type='Lookup'>" + int.Parse(sourceItem[WebHookList.CostCenter]) + "</Value></Eq>"
                        //                                                + "<Eq><FieldRef Name='GL_x0020_Account'  LookupId='TRUE'/><Value Type='Lookup'>" + int.Parse(sourceItem[WebHookList.GLAccount]) + "</Value></Eq>"
                        //                                            + "</And>"
                        //                                   + "</And>"
                        //                                + "</And>"
                        //                             + "</Where></Query></View>";

                        //    ListItemCollection targetItems = this.sharePointContext.Web.Lists.GetByTitle(ExpenseSummaryList.DefaultListTitle).GetItems(targetQuery);
                        //    this.sharePointContext.Load(targetItems);
                        //    this.sharePointContext.ExecuteQuery();
                        //    ListItem oListItem;
                        //    if (targetItems.Count == 0)
                        //    {
                        //        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        //        oListItem = this.sharePointContext.Web.Lists.GetByTitle(ExpenseSummaryList.DefaultListTitle).AddItem(itemCreateInfo);

                        //        oListItem[ExpenseSummaryList.ChargedCompany] = CompanyLookup;
                        //        oListItem[ExpenseSummaryList.CostCenter] = CostCenterLookup;
                        //        oListItem[ExpenseSummaryList.GLAccount] = GLAccountLookup;
                        //        oListItem[ExpenseSummaryList.FiscalYear] = sourceItem[WebHookList.FiscalYear];
                        //    }
                        //    else
                        //    {
                        //        oListItem = targetItems[0];
                        //    }
                        //    this.log.Info("ColName:" + ExpenseSummaryList.JanAmt + sourceItem["Phase"]);
                        //    this.log.Info("JanAmt:" + sourceItem[WebHookList.JanAmt]);
                        //    oListItem[ExpenseSummaryList.JanAmt + sourceItem["Phase"]] = sourceItem[WebHookList.JanAmt];
                        //    oListItem[ExpenseSummaryList.FebAmt + sourceItem["Phase"]] = sourceItem[WebHookList.FebAmt];
                        //    oListItem[ExpenseSummaryList.MarAmt + sourceItem["Phase"]] = sourceItem[WebHookList.MarAmt];
                        //    oListItem[ExpenseSummaryList.AprAmt + sourceItem["Phase"]] = sourceItem[WebHookList.AprAmt];
                        //    oListItem[ExpenseSummaryList.MayAmt + sourceItem["Phase"]] = sourceItem[WebHookList.MayAmt];
                        //    oListItem[ExpenseSummaryList.JunAmt + sourceItem["Phase"]] = sourceItem[WebHookList.JunAmt];
                        //    oListItem[ExpenseSummaryList.JulAmt + sourceItem["Phase"]] = sourceItem[WebHookList.JulAmt];
                        //    oListItem[ExpenseSummaryList.AugAmt + sourceItem["Phase"]] = sourceItem[WebHookList.AugAmt];
                        //    oListItem[ExpenseSummaryList.SepAmt + sourceItem["Phase"]] = sourceItem[WebHookList.SepAmt];
                        //    oListItem[ExpenseSummaryList.OctAmt + sourceItem["Phase"]] = sourceItem[WebHookList.OctAmt];
                        //    oListItem[ExpenseSummaryList.NovAmt + sourceItem["Phase"]] = sourceItem[WebHookList.NovAmt];
                        //    oListItem[ExpenseSummaryList.DecAmt + sourceItem["Phase"]] = sourceItem[WebHookList.DecAmt];
                        //    if (sourceItem["Phase"] != "Estimate")
                        //    {
                        //        oListItem[ExpenseSummaryList.FYNext1 + sourceItem["Phase"]] = sourceItem[WebHookList.FYNext1];
                        //        oListItem[ExpenseSummaryList.FYNext2 + sourceItem["Phase"]] = sourceItem[WebHookList.FYNext2];
                        //    }

                        //    oListItem.Update();
                        #endregion
                        this.sharePointContext.ExecuteQuery();
                        itemsProcessed++;
                    }

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
