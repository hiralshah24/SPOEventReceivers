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
        public const String WebhookProcessedDate = "WebhookProcessed";
        const String Created = "Created";

        public WebHookList(ClientContext spContext) : base(spContext)
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
                CamlQuery itemsQuery = CamlQuery.CreateAllItemsQuery(1, "Id", "Title", "Created", "Author", "Modified", "ContentTypeId", ChargedCompany, CostCenter, FiscalYear, EventCategory, WebhookProcessedDate);
                this.LoadList(itemsQuery);
            }

            List<ListItem> returnItems = new List<ListItem>();

            foreach (Int32 currentItemId in itemIdsToBeProcessed)
            {
                try
                {
                    ListItem eventItem = this.List.GetItemById(currentItemId);
                    this.SharePointContext.Load(eventItem, item => item.Id, item => item["Title"], item => item["Created"], item => item["Author"], item => item["Modified"], item => item["ContentTypeId"]
                                                                                , item => item[WebhookProcessedDate], item => item[CostCenter], item => item[ChargedCompany], item => item[EventCategory]
                                                                                , item => item[GLAccount], item => item[JanAmt], item => item[FebAmt], item => item[MarAmt]
                                                                                , item => item[AprAmt], item => item[MayAmt], item => item[JunAmt], item => item[JunAmt]
                                                                                , item => item[AugAmt], item => item[SepAmt], item => item[OctAmt], item => item[NovAmt] 
                                                                                , item => item[DecAmt], item => item[FYNext1], item => item[FYNext2], item => item[FiscalYear]);

                        this.SharePointContext.ExecuteQuery();
                    if (eventItem[EventCategory].ToString() == "IT")
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
            
            //Get Cost Center by GL Account
            SPList findCostCenterCodeList = web.Lists["Cost Centers"];
            string key = itemprop["costcenterKey"];
            SPListItem costCenterCodeItem = findCostCenterCodeList.GetItemById(Convert.ToInt32(key.Substring(0, key.IndexOf(";"))));
            string costCenterCode = costCenterCodeItem["Code"].ToString();

            //get Company code by Owning comapny
            SPList findCompanyCodeList = web.Lists["Companies"];
            string coKey = itemprop["CompanyID"];
            SPListItem companyCodeItem = findCompanyCodeList.GetItemById(Convert.ToInt32(coKey.Substring(0, coKey.IndexOf(";"))));
            string companyCode = companyCodeItem["Code"].ToString();
            SPList listToQuery = web.Lists[listTitle];
            SPQuery query = new SPQuery();
            //Get All Items by Cost Center and Company
            query.Folder = listToQuery.RootFolder.SubFolders[itemprop["fiscalYear"].ToString()].SubFolders[companyCode + "-" + costCenterCode];
            //query.Query = "<Where><Eq><FieldRef Name='GL_x0020_Account'/><Value Type='Lookup'>" + trimGL + "</Value></Eq></Where>";
            query.Query = "<Where>"
                            + "<And>"
                                + "<Eq><FieldRef Name='GL_x0020_Account'/><Value Type='Lookup'>" + GLAccountLookup.LookupValue + "</Value></Eq>"
                                + "<Eq><FieldRef Name='Expense_Category' /><Value Type='Choice'>IT</Value></Eq>"
                            + "</And>"
                         + "</Where>";
            query.ViewAttributes = "Scope=\"Recursive\"";


            throw new NotImplementedException();
        }
    }
}
