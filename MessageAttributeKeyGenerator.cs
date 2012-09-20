using System;
using Microsoft.Office.Interop.Outlook;

namespace com.mattberther.deduper
{
    class MessageAttributeKeyGenerator : IKeyGenerator
    {
        public string CreateKey(MailItem item)
        {
            return String.Format("{0} {1} {2:yyyyMMddhhmmss}",
                  item.SenderEmailAddress, item.Subject, item.SentOn);
        }
    }
}
