using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using Microsoft.Azure.WebJobs.Host;
namespace AzureFunctions
{
    public class WebHookList : SharePointList
    {
        new internal static String DefaultListTitle = "Expense Details Preliminary";

        const String Title = "Title";
        public String Phase = "Preliminary";
        public const String ChargedCompany = "Owning_x0020_Company";
        public const String CostCenter = "Owning_x0020_CostCenter";
        public const String FiscalYear = "Fiscal_Year";
        public const String ExpenseCategory = "Expense_Category";
        public const String GLAccount = "GL_x0020_Account";
        public const String JanAmt = "Jan_Amount";
        public const String FebAmt = "Feb_Amount";
        public const String MarAmt = "Mar_Amount";
        public const String AprAmt = "Apr_Amount";
        public const String MayAmt = "May_Amount";
        public const String JunAmt = "Jun_Amount";
        public const String JulAmt = "Jul_Amount";
        public const String AugAmt = "Aug_Amount";
        public const String SepAmt = "Sep_Amount";
        public const String OctAmt = "Oct_Amount";
        public const String NovAmt = "Nov_Amount";
        public const String DecAmt = "Dec_Amount";
        public const String FYNext1 = "FY_x002b_1_Plan_Amount";
        public const String FYNext2 = "FY_x002b_2_Plan_Amount";
        public const String WebhookProcessedDate = "WebhookProcessed";
        public const String CompanyCode = "Owning_x0020_Company_x003a_Code";
        public const String CostCenterCode = "Charged_x0020_CostCenter_x003a_Code";
        const String Created = "Created";

        public WebHookList(ClientContext spContext, Guid listId) : base(spContext)
        {
            this.ListTitle = DefaultListTitle;
            IEnumerable<List> eventListQueryResults = spContext.LoadQuery<List>(spContext.Web.Lists.Where(lst => lst.Id == listId));
            spContext.ExecuteQueryRetry();
            switch (eventListQueryResults.FirstOrDefault().Title)
            {
                case "Expense Details Actuals":
                    this.ListTitle = "Expense Details Actuals";                    
                    Phase = "Actual";
                    break;
                case "Expense Details Final":
                    this.ListTitle = "Expense Details Final";
                    Phase = "Final";
                    break;
                case "Expense Details Estimates":
                    this.ListTitle = "Expense Details Estimates";
                    Phase = "Estimate";
                    break;
                case "Expense Details Preliminary":
                    this.ListTitle = "Expense Details Preliminary";
                    Phase = "Preliminary";
                    break;
                default:
                    break;
            }
            DefaultListTitle = this.ListTitle;
        }

        /// <summary>
        /// Given a <see cref="List"/> of item ids, return a <see cref="List"/> of <see cref="ListItem"/>s 
        /// corresponding to those Ids.
        /// </summary>
        /// <param name="itemIdsToBeProcessed"></param>
        public override List<ListItem> GetItemsById(HashSet<Int32> itemIdsToBeProcessed)
        {
            this.List = this.SharePointContext.Web.Lists.GetByTitle(WebHookList.DefaultListTitle);
            this.SharePointContext.Load(this.List);
            this.SharePointContext.ExecuteQuery();
            if (this.List == null)
            {
                //  Using this method call to initialize this.List.  Completed data is retrieved below.  
                CamlQuery itemsQuery = CamlQuery.CreateAllItemsQuery(1, "Id", "Title", "Created", "Author", "Modified", "ContentTypeId", ChargedCompany, CostCenter, FiscalYear, ExpenseCategory, WebhookProcessedDate);
                this.LoadList(itemsQuery);
            }

            List<ListItem> returnItems = new List<ListItem>();

            foreach (Int32 currentItemId in itemIdsToBeProcessed)
            {
                try
                {
                    ListItem eventItem = this.List.GetItemById(currentItemId);
                    this.SharePointContext.Load(eventItem, item => item.Id, item => item["Title"], item => item["Created"], item => item["Author"], item => item["Modified"], item => item["ContentTypeId"]
                                                                                , item => item[WebhookProcessedDate], item => item[CostCenter], item => item[ChargedCompany], item => item[ExpenseCategory]
                                                                                , item => item[GLAccount], item => item[JanAmt], item => item[FebAmt], item => item[MarAmt]
                                                                                , item => item[AprAmt], item => item[MayAmt], item => item[JunAmt], item => item[JunAmt]
                                                                                , item => item[AugAmt], item => item[SepAmt], item => item[OctAmt], item => item[NovAmt]
                                                                                , item => item[DecAmt], item => item[FYNext1], item => item[FYNext2], item => item[FiscalYear]);//, item => item[CompanyCode], item => item[CostCenterCode]);

                    this.SharePointContext.ExecuteQuery();
                    if (eventItem[ExpenseCategory].ToString() == "IT")
                    {
                        returnItems.Add(eventItem);
                    }
                }
                catch (Microsoft.SharePoint.Client.ServerException)
                {

                }
                catch (Exception ex)
                {

                }
            }

            return returnItems;
        }

        internal Dictionary<string, string> ProcessItem(ListItem currentItem)
        {
            FieldLookupValue GLAccountLookup = currentItem[GLAccount] as FieldLookupValue;
            FieldLookupValue CompanyLookup = currentItem[ChargedCompany] as FieldLookupValue;
            FieldLookupValue CostCenterLookup = currentItem[CostCenter] as FieldLookupValue;
            //FieldLookupValue CompanyCodeV = currentItem[CompanyCode] as FieldLookupValue;
            Dictionary<string, string> SourceData = new Dictionary<string, string>();
            Folder folder = this.SharePointContext.Web.GetFolderByServerRelativeUrl(this.SharePointContext.Url + "/Lists/" + 
                            WebHookList.DefaultListTitle + "/" + currentItem[WebHookList.FiscalYear] + "/" +
                            CompanyLookup.LookupValue.Split('-')[0].Trim() + "-" + CostCenterLookup.LookupValue.Split('-')[0].Trim());
            this.SharePointContext.Load(folder);
            this.SharePointContext.ExecuteQuery();
            CamlQuery query = new CamlQuery();
            query.FolderServerRelativeUrl = folder.ServerRelativeUrl;
            
            query.ViewXml = "<View Scope=\"RecursiveAll\"> " +
                                "<Query>" +
                                    "<Where>"
                                    + "<And>"
                                        + "<Eq><FieldRef Name='GL_x0020_Account'/><Value Type='Lookup'>" + GLAccountLookup.LookupValue + "</Value></Eq>"
                                        + "<Eq><FieldRef Name='Expense_Category' /><Value Type='Choice'>IT</Value></Eq>"
                                    + "</And>"
                                    + "</Where>"
                                + "</Query>"
                            + "</View>";
            ListItemCollection colItems = this.SharePointContext.Web.Lists.GetByTitle(WebHookList.DefaultListTitle).GetItems(query);
            this.SharePointContext.Load(colItems);
            this.SharePointContext.ExecuteQuery();
            Int64 jan = 0;
            Int64 feb = 0;
            Int64 mar = 0;
            Int64 apr = 0;
            Int64 may = 0;
            Int64 jun = 0;
            Int64 jul = 0;
            Int64 aug = 0;
            Int64 sep = 0;
            Int64 oct = 0;
            Int64 nov = 0;
            Int64 dec = 0;
            Int64 fynext = 0;
            Int64 fynextnext = 0;
            if (colItems != null && colItems.Count > 0)
            {
                foreach (ListItem colItem in colItems)
                {
                    jan += Convert.ToInt64(colItem[JanAmt]);
                    feb += Convert.ToInt64(colItem[FebAmt]);
                    mar += Convert.ToInt64(colItem[MarAmt]);
                    apr += Convert.ToInt64(colItem[AprAmt]);
                    may += Convert.ToInt64(colItem[MayAmt]);
                    jun += Convert.ToInt64(colItem[JunAmt]);
                    jul += Convert.ToInt64(colItem[JulAmt]);
                    aug += Convert.ToInt64(colItem[AugAmt]);
                    sep += Convert.ToInt64(colItem[SepAmt]);
                    oct += Convert.ToInt64(colItem[OctAmt]);
                    nov += Convert.ToInt64(colItem[NovAmt]);
                    dec += Convert.ToInt64(colItem[DecAmt]);
                    fynext += Convert.ToInt64(colItem[FYNext1]);
                    fynextnext += Convert.ToInt64(colItem[FYNext2]);
                }
                
                SourceData.Add("Phase", Phase);
                SourceData.Add(WebHookList.ChargedCompany, CompanyLookup.LookupId.ToString());
                SourceData.Add(WebHookList.CostCenter, CostCenterLookup.LookupId.ToString());
                SourceData.Add(WebHookList.GLAccount, GLAccountLookup.LookupId.ToString());
                SourceData.Add(WebHookList.FiscalYear, currentItem[FiscalYear].ToString());
                SourceData.Add(WebHookList.JanAmt, jan.ToString());
                SourceData.Add(WebHookList.FebAmt, feb.ToString());
                SourceData.Add(WebHookList.MarAmt, mar.ToString());
                SourceData.Add(WebHookList.AprAmt, apr.ToString());
                SourceData.Add(WebHookList.MayAmt, may.ToString());
                SourceData.Add(WebHookList.JunAmt, jun.ToString());
                SourceData.Add(WebHookList.JulAmt, jul.ToString());
                SourceData.Add(WebHookList.AugAmt, aug.ToString());
                SourceData.Add(WebHookList.SepAmt, sep.ToString());
                SourceData.Add(WebHookList.OctAmt, oct.ToString());
                SourceData.Add(WebHookList.NovAmt, nov.ToString());
                SourceData.Add(WebHookList.DecAmt, dec.ToString());
                SourceData.Add(WebHookList.FYNext1, fynext.ToString());
                SourceData.Add(WebHookList.FYNext2, fynextnext.ToString());
            }
            return SourceData;
        }
    }
}
