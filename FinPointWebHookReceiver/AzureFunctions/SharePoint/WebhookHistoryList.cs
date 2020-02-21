using Microsoft.SharePoint.Client;
using System;
using System.Linq;

namespace AzureFunctions
{
    public class WebhookHistoryList : SharePointList
    {
        new internal static String DefaultListTitle = "Webhook History";

        const String ProcessName = "Title";
        const String ItemsReceived = "ProcessorItemsReceived";
        const String ItemsProcessed = "ProcessorItemsProcessed";
        const String StartTime = "ProcessorStart";
        const String EndTime = "ProcessorEnd";
        const String EndingChangeToken = "EndingChangeToken";
        const String Created = "Created";

        public WebhookHistoryList(ClientContext spContext) : base(spContext)
        {
            this.ListTitle = DefaultListTitle;
        }

        public String GetLastToken(string processName)
        {
            String returnValue = String.Empty;

            if (this.List == null)
            {
                CamlQuery loadQuery = CamlQuery.CreateAllItemsQuery(20000, "Id", "Title", "Created", "Author", "Modified", "ContentTypeId", ProcessName, ItemsReceived, ItemsProcessed
                                                                    , StartTime, EndTime, EndingChangeToken);
                loadQuery.ViewXml = loadQuery.ViewXml.Replace("<ViewFields>", String.Format("<Query><Where><Eq><FieldRef Name='{0}'/><Value Type='Text'>{1}</Value></Eq></Where><OrderBy><FieldRef Name='{2}' Ascending='False'/></OrderBy></Query><RowLimit>1</RowLimit><ViewFields>", ProcessName, processName, Created));
                this.LoadList(loadQuery);
            }

            var historyItems = from item in this.ListItems
                               select item;

            try
            {
                //  If the query returns no rows, count blows up with this exception.  Haven't figured out another 
                //   way to check for a "hit".
                if (historyItems.Count() > 0)
                {
                    returnValue = Convert.ToString(historyItems.First()[EndingChangeToken]);
                }
            }
            catch (PropertyOrFieldNotInitializedException)
            {
            }

            return returnValue;
        }

        public Boolean TryAddItem(String processName, Int32 itemsReceived, Int32 itemsProcessed, DateTime startTime, DateTime endTime, String lastChangeToken)
        {
            Boolean returnValue = true;

            if (this.List == null)
            {
                this.LoadList(false);
            }

            ListItemCreationInformation itemCreateInfo = new ListItemCreationInformation();
            ListItem oListItem = this.List.AddItem(itemCreateInfo);
            oListItem[ProcessName] = processName;
            oListItem[ItemsReceived] = itemsReceived;
            oListItem[ItemsProcessed] = itemsProcessed;
            oListItem[StartTime] = startTime;
            oListItem[EndTime] = endTime;
            oListItem[EndingChangeToken] = lastChangeToken;
            oListItem.Update();
            this.SharePointContext.ExecuteQuery();

            return returnValue;
        }

    }
}
