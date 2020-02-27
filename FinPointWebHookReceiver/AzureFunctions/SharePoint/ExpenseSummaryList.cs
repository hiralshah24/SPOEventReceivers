using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;

namespace AzureFunctions
{
    public class ExpenseSummaryList : SharePointList
    {
        new internal static String DefaultListTitle = "Expense Summary_CC_GL";

        const String ProcessName = "Title";
        public const String ChargedCompany = "Owning_x0020_Company";
        public const String CostCenter = "Owning_x0020_CostCenter";
        public const String FiscalYear = "Fiscal_Year";
        public const String ExpenseCategory = "Expense_Category";
        public const String GLAccount = "GL_x0020_Account";
        public const String JanAmt = "Jan_";
        public const String FebAmt = "Feb_";
        public const String MarAmt = "Mar_";
        public const String AprAmt = "Apr_";
        public const String MayAmt = "May_";
        public const String JunAmt = "Jun_";
        public const String JulAmt = "Jul_";
        public const String AugAmt = "Aug_";
        public const String SepAmt = "Sep_";
        public const String OctAmt = "Oct_";
        public const String NovAmt = "Nov_";
        public const String DecAmt = "Dec_";
        public const String FYNext1 = "FY_x002b_1_Plan_";
        public const String FYNext2 = "FY_x002b_2_Plan_";

        public ExpenseSummaryList(ClientContext spContext) : base(spContext)
        {
            this.ListTitle = DefaultListTitle;
        }


        /// <summary>
        /// Given a <see cref="List"/> of item ids, return a <see cref="List"/> of <see cref="ListItem"/>s 
        /// corresponding to those Ids.
        /// </summary>
        /// <param name="itemIdsToBeProcessed"></param>
        public override List<ListItem> GetItemsById(HashSet<Int32> itemIdsToBeProcessed)
        {
            if (this.List == null)
            {
                //  Using this method call to initialize this.List.  Completed data is retrieved below.
                CamlQuery itemsQuery = CamlQuery.CreateAllItemsQuery(20000, "Id", "Title", "Created", "Author", "Modified", "ContentTypeId", ProcessName, ChargedCompany, CostCenter, FiscalYear);
                //CamlQuery itemsQuery = CamlQuery.CreateAllItemsQuery(1, "Id", "Title", "Created", "Author", "Modified", "ContentTypeId", ChargedCompany, CostCenter, FiscalYear, WebhookProcessedDate);
                this.LoadList(itemsQuery);
            }

            List<ListItem> returnItems = new List<ListItem>();

            foreach (Int32 currentItemId in itemIdsToBeProcessed)
            {
                ListItem eventItem = this.List.GetItemById(currentItemId);
                this.SharePointContext.Load(eventItem, item => item.Id, item => item["Title"], item => item["Created"], item => item["Author"], item => item["Modified"], item => item["ContentTypeId"]
                                                                                , item => item[CostCenter], item => item[ChargedCompany], item => item[FiscalYear]);
                try
                {
                    this.SharePointContext.ExecuteQuery();
                    returnItems.Add(eventItem);
                }
                catch (Microsoft.SharePoint.Client.ServerException)
                {
                }
            }

            return returnItems;
        }
        
        internal bool TryAddUpdateItem(Dictionary<string, string> sourceItem)
        {
            Boolean returnValue = false;
            if (this.List == null)
            {
                this.LoadList(false);
            }
            if (sourceItem != null && sourceItem.Count > 0)
            {
                try
                {
                    FieldLookupValue CompanyLookup = new FieldLookupValue();
                    CompanyLookup.LookupId = int.Parse(sourceItem[WebHookList.ChargedCompany]);
                    FieldLookupValue CostCenterLookup = new FieldLookupValue();
                    CostCenterLookup.LookupId = int.Parse(sourceItem[WebHookList.CostCenter]);
                    FieldLookupValue GLAccountLookup = new FieldLookupValue();
                    GLAccountLookup.LookupId = int.Parse(sourceItem[WebHookList.GLAccount]);
                    CamlQuery targetQuery = new CamlQuery();

                    targetQuery.ViewXml = "<View>"
                                            + "<Query>"
                                            + "<Where>"
                                                + "<And>"
                                                    + "<Eq><FieldRef Name='Owning_x0020_Company' LookupId='TRUE' /><Value Type='Lookup'>" + int.Parse(sourceItem[ExpenseSummaryList.ChargedCompany]) + "</Value></Eq>"
                                                    + "<And>"
                                                        + "<Eq><FieldRef Name='Fiscal_Year'/><Value Type='Number'>" + int.Parse(sourceItem[ExpenseSummaryList.FiscalYear]) + "</Value></Eq>"
                                                            + "<And>"
                                                                + "<Eq><FieldRef Name='Owning_x0020_CostCenter' LookupId='TRUE'/><Value Type='Lookup'>" + int.Parse(sourceItem[ExpenseSummaryList.CostCenter]) + "</Value></Eq>"
                                                                + "<Eq><FieldRef Name='GL_x0020_Account'  LookupId='TRUE'/><Value Type='Lookup'>" + int.Parse(sourceItem[ExpenseSummaryList.GLAccount]) + "</Value></Eq>"
                                                            + "</And>"
                                                   + "</And>"
                                                + "</And>"
                                             + "</Where></Query></View>";

                    ListItemCollection targetItems = this.List.GetItems(targetQuery);
                    this.SharePointContext.Load(targetItems);
                    this.SharePointContext.ExecuteQuery();
                    ListItem oListItem;
                    if (targetItems.Count == 0)
                    {
                        ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
                        oListItem = this.List.AddItem(itemCreateInfo);

                        oListItem[ChargedCompany] = CompanyLookup;
                        oListItem[CostCenter] = CostCenterLookup;
                        oListItem[GLAccount] = GLAccountLookup;
                        oListItem[FiscalYear] = sourceItem[WebHookList.FiscalYear];
                    }
                    else
                    {
                        oListItem = targetItems[0];
                    }

                    oListItem[JanAmt + sourceItem["Phase"]] = sourceItem[WebHookList.JanAmt];
                    oListItem[FebAmt + sourceItem["Phase"]] = sourceItem[WebHookList.FebAmt];
                    oListItem[MarAmt + sourceItem["Phase"]] = sourceItem[WebHookList.MarAmt];
                    oListItem[AprAmt + sourceItem["Phase"]] = sourceItem[WebHookList.AprAmt];
                    oListItem[MayAmt + sourceItem["Phase"]] = sourceItem[WebHookList.MayAmt];
                    oListItem[JunAmt + sourceItem["Phase"]] = sourceItem[WebHookList.JunAmt];
                    oListItem[JulAmt + sourceItem["Phase"]] = sourceItem[WebHookList.JulAmt];
                    oListItem[AugAmt + sourceItem["Phase"]] = sourceItem[WebHookList.AugAmt];
                    oListItem[SepAmt + sourceItem["Phase"]] = sourceItem[WebHookList.SepAmt];
                    oListItem[OctAmt + sourceItem["Phase"]] = sourceItem[WebHookList.OctAmt];
                    oListItem[NovAmt + sourceItem["Phase"]] = sourceItem[WebHookList.NovAmt];
                    oListItem[DecAmt + sourceItem["Phase"]] = sourceItem[WebHookList.DecAmt];
                    if (sourceItem["Phase"] != "Estimate")
                    {
                        oListItem[FYNext1 + sourceItem["Phase"]] = sourceItem[WebHookList.FYNext1];
                        oListItem[FYNext2 + sourceItem["Phase"]] = sourceItem[WebHookList.FYNext2];
                    }
                    oListItem.Update();
                    this.SharePointContext.ExecuteQuery();
                }
                catch (Exception ex) { }
            }



            return returnValue;
        }


    }
}
