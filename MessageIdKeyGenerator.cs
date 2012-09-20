using Microsoft.Office.Interop.Outlook;

namespace com.mattberther.deduper
{
    class MessageIdKeyGenerator : IKeyGenerator
    {
        public string CreateKey(MailItem item)
        {
            const string internetMessageIdWTag = "http://schemas.microsoft.com/mapi/proptag/0x1035001F";
            return (string)item.PropertyAccessor.GetProperty(internetMessageIdWTag);
        }
    }
}
