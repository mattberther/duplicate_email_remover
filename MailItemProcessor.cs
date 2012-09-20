using System.Collections.Generic;
using Microsoft.Office.Interop.Outlook;

namespace com.mattberther.deduper
{
    class MailItemProcessor
    {
        private readonly List<string> messageKeys = new List<string>();
        private readonly List<IKeyGenerator> keyGenerators = new List<IKeyGenerator>();
        
        public void RegisterKeyGenerator(IKeyGenerator keyGenerator)
        {
            keyGenerators.Add(keyGenerator);
        }

        public bool DryRun { get; set; }
        public int ItemsDeleted { get; private set; }

        public void ProcessMailItem(MailItem item)
        {
            var markedForDeletion = false;

            keyGenerators.ForEach(delegate(IKeyGenerator generator)
                                    {
                                        var key = generator.CreateKey(item);
                                        if (IsDuplicateMessage(key))
                                        {
                                            markedForDeletion = true;
                                        }
                                        else
                                        {
                                            messageKeys.Add(key);
                                        }
                                    });

            if (markedForDeletion)
            {
                if (!DryRun) { item.Delete(); }
                ItemsDeleted++;
            }
        }

        private bool IsDuplicateMessage(string key)
        {
            return messageKeys.Contains(key);
        }
    }
}
