using System;
using Microsoft.Office.Interop.Outlook;

namespace com.mattberther.deduper
{
    class Program
    {
        static void Main(string[] args)
        {
            var mailItemProcessor = new MailItemProcessor(true);
            mailItemProcessor.RegisterKeyGenerator(new MessageIdKeyGenerator());
            mailItemProcessor.RegisterKeyGenerator(new MessageAttributeKeyGenerator());

            var outlook = new Application();
            var outlookNamespace = outlook.GetNamespace("MAPI");
            var folder = outlookNamespace.Folders[1].Folders["Archive"];

            foreach (var item in folder.Items)
            {
                var mailItem = item as MailItem;
                if (mailItem != null)
                {
                    mailItemProcessor.ProcessMailItem(mailItem);    
                }
            }

            Console.WriteLine("Deleted items: {0}", mailItemProcessor.ItemsDeleted);
            Console.WriteLine("Total items: {0}", folder.Items.Count);
            Console.WriteLine("** PRESS ENTER **");
            Console.ReadLine();
        }
    }
}
