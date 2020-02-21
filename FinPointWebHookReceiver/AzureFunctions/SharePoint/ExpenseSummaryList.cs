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
        public const String EventCategory = "Expense_Category";
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
        public Boolean TryAddItem(int CompanyId, int costCenterId, String fiscalYear)
        {
            Boolean returnValue = true;

            if (this.List == null)
            {
                this.LoadList(false);
            }

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem oListItem = this.List.AddItem(itemCreateInfo);
            FieldLookupValue CompanyLookup = new FieldLookupValue();
            CompanyLookup.LookupId = CompanyId;
            FieldLookupValue CostCenterLookup = new FieldLookupValue();
            CostCenterLookup.LookupId = costCenterId;

            oListItem[ChargedCompany] = CompanyLookup;
            oListItem[CostCenter] = CostCenterLookup;
            oListItem[FiscalYear] = fiscalYear;
            
            oListItem.Update();
            this.SharePointContext.ExecuteQuery();

            return returnValue;
        }
        internal bool TryAddItem(ListItem currentItem)
        {
            Boolean returnValue = true;
            if (this.List == null)
            {
                this.LoadList(false);
            }
            ListItem sourceItem = currentItem;
            


            FieldLookupValue companyLookup = sourceItem[WebHookList.ChargedCompany] as FieldLookupValue;
            FieldLookupValue costCenterLookup = sourceItem[WebHookList.CostCenter] as FieldLookupValue;
            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem oListItem = this.List.AddItem(itemCreateInfo);
            FieldLookupValue CompanyLookup = new FieldLookupValue();
            CompanyLookup.LookupId = companyLookup.LookupId;
            FieldLookupValue CostCenterLookup = new FieldLookupValue();
            CostCenterLookup.LookupId = costCenterLookup.LookupId;

            oListItem[ChargedCompany] = CompanyLookup;
            oListItem[CostCenter] = CostCenterLookup;
            oListItem[FiscalYear] = sourceItem[WebHookList.FiscalYear];

            oListItem.Update();
            this.SharePointContext.ExecuteQuery();

            return returnValue;
        }
    }
}
