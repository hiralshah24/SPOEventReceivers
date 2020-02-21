using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;

namespace AzureFunctions
{
    public class SharePointList
    {
        public static String DefaultListTitle = "Expense Details Preliminary";
        protected ClientContext SharePointContext;
        protected List List;
        protected ListItemCollection ListItems;
        String listTitle = String.Empty;

        public String ListTitle
        {
            set
            {
                this.listTitle = value;
            }
            get
            {
                return this.listTitle;
            }
        }

        protected SharePointList(ClientContext spContext)
        {
            this.SharePointContext = spContext;
            this.ListTitle = DefaultListTitle;
        }

        protected void LoadList(Boolean loadItems)
        {
            this.List = this.SharePointContext.Web.Lists.GetByTitle(this.ListTitle);
            this.SharePointContext.Load(this.List.Fields, fields => fields.Include(field => field.Id, field => field.Title, field => field.StaticName, field => field.InternalName, field => field.TypeDisplayName, Field => Field.FieldTypeKind, field => field.TypeAsString));
            this.SharePointContext.Load(this.List);

            if (loadItems)
            {
                this.ListItems = this.List.GetItems(CamlQuery.CreateAllItemsQuery(20000));
                this.SharePointContext.Load(this.ListItems);

                /*  Parent List Title and Fields are required to do token substitution.  */
                //  TODO:  Optimize:  Only load these for lists for which we are going to use them?
                this.SharePointContext.Load(this.ListItems, items => items.Include(item => item.ParentList.Title));
                this.SharePointContext.Load(this.ListItems, items => items.Include(item => item.ParentList.Fields));
                this.SharePointContext.Load(this.ListItems, items => items.Include(item => item.ContentType.Fields));
            }

            this.SharePointContext.ExecuteQuery();
        }

        protected void LoadList(CamlQuery loadQuery)
        {
            this.List = this.SharePointContext.Web.Lists.GetByTitle(this.ListTitle);
            this.SharePointContext.Load(this.List);
            this.SharePointContext.Load(this.List.Fields, fields => fields.Include(field => field.Id, field => field.Title, field => field.StaticName, field => field.InternalName, field => field.TypeDisplayName, Field => Field.FieldTypeKind, field => field.TypeAsString));
            this.ListItems = this.List.GetItems(loadQuery);
            this.SharePointContext.Load(this.ListItems);

            /*  Parent List Title and Fields are required to do token substitution.  */
            //  TODO:  Optimize:  Only load these for lists for which we are going to use them?
            this.SharePointContext.Load(this.ListItems, items => items.Include(item => item.ParentList.Title));
            this.SharePointContext.Load(this.ListItems, items => items.Include(item => item.ParentList.Fields));
            this.SharePointContext.Load(this.ListItems, items => items.Include(item => item.ContentType.Fields));

            this.SharePointContext.ExecuteQuery();
        }

        /// <summary>
        /// Given a <see cref="List"/> of item ids, return a <see cref="List"/> of <see cref="ListItem"/>s 
        /// corresponding to those Ids.
        /// </summary>
        /// <param name="itemIdsToBeProcessed"></param>
        public virtual List<ListItem> GetItemsById(HashSet<Int32> itemIdsToBeProcessed)
        {
            if (this.List == null)
            {
                CamlQuery itemsQuery = CamlQuery.CreateAllItemsQuery(1, "Id", "Title", "Created", "Author", "Modified", "ContentTypeId");

                this.LoadList(itemsQuery);
            }

            List<ListItem> returnItems = new List<ListItem>();

            foreach (Int32 currentItemId in itemIdsToBeProcessed)
            {
                ListItem eventItem = this.List.GetItemById(currentItemId);
                this.SharePointContext.Load(eventItem, item => item.Id, item => item["Title"], item => item["Created"], item => item["Author"], item => item["Modified"]
                                                            , item => item["ContentTypeId"], item => item.ParentList.Title);
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

    }
}
