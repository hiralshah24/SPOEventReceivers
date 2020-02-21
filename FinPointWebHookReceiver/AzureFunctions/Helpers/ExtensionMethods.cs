using Microsoft.SharePoint.Client;
using System;

namespace AzureFunctions
{
    public static class ExtensionMethods
    {
        public static String GetVersion(this ChangeToken ct)
        {
            return ct.StringValue.Split(';')[0];
        }

        public static String GetChangeScope(this ChangeToken ct)
        {
            return ct.StringValue.Split(';')[1];
        }

        public static String GetScopeId(this ChangeToken ct)
        {
            return ct.StringValue.Split(';')[2];
        }

        public static String GetDateAsString(this ChangeToken ct)
        {
            string ticks = ct.StringValue.Split(';')[3];
            DateTime dt = new DateTime(Convert.ToInt64(ticks));

            return string.Format("{0} {1}", dt.ToShortDateString(), dt.ToLongTimeString());
        }

        public static DateTime GetDate(this ChangeToken ct)
        {
            string ticks = ct.StringValue.Split(';')[3];
            DateTime dt = new DateTime(Convert.ToInt64(ticks));

            return dt;
        }

        public static String GetChangeNumber(this ChangeToken ct)
        {
            return ct.StringValue.Split(';')[4];
        }
    }
}
