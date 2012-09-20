using Microsoft.Office.Interop.Outlook;

namespace com.mattberther.deduper
{
    internal interface IKeyGenerator
    {
        string CreateKey(MailItem item);
    }
}