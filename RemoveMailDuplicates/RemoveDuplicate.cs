using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;

namespace RemoveMailDuplicates
{
    class RemoveDuplicate
    {
        private class Mail
        {
            public MailItem Item { get; set; }
            public string Subject { get; set; }
            public string Body { get; set; }
            public int AttachmentCount { get; set; }
            public DateTime CreationTime { get; set; }
            public long Size { get; set; }

            public Mail(MailItem item)
            {
                Item = item;
                Subject = item.Subject;
                Body = item.Body;
                AttachmentCount = item.Attachments.Count;
                CreationTime = item.CreationTime;
                Size = item.Size;
            }
        }

        internal static async Task ProcessFolder(Application app, MAPIFolder currentFolder, bool recursive)
        {
            Debug.Print(currentFolder.FullFolderPath);
            var agg = new Dictionary<string, List<Mail>>();
            foreach (var item in currentFolder.Items)
            {
                var mail = item as MailItem;
                if (mail == null || mail.ConversationID == null)
                {
                    continue;
                }
                List<Mail> list;
                var id = mail.CreationTime.ToString("yyyyMMddhhmmss") + mail.SenderName + mail.Subject;
                //var id = mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x1035001F") + mail.Subject;
                 
                if (!agg.TryGetValue(id, out list))
                {
                    list = new List<Mail>();
                    agg.Add(id, list);
                }
                list.Add(new Mail(mail));
            }

            foreach (var pair in agg)
            {
                if (pair.Value.Count > 1)
                {
                    Debug.Print("-----");
                    Mail etalon = null;
                    foreach (var item in pair.Value)
                    {
                        Debug.Print(item.Subject);
                        if (etalon == null)
                        {
                            etalon = item;
                        }
                        else {
                            if (etalon.Subject == item.Subject
                                && etalon.AttachmentCount >= item.AttachmentCount
                                && ( etalon.Body == item.Body || string.IsNullOrWhiteSpace(item.Body) )
                                && etalon.CreationTime == item.CreationTime 
                                && etalon.Size >= item.Size)
                            {
                                Debug.Print("Delete");
                                item.Item.Delete();
                            }
                            else
                            {
                                if ( (etalon.AttachmentCount == 0 && item.AttachmentCount > 0)
                                    || 
                                    (string.IsNullOrWhiteSpace(etalon.Body))
                                    )
                                {
                                    Debug.Print("Delete");
                                    etalon.Item.Delete();
                                    etalon = item;
                                } else
                                {
                                    Debug.Print("Something wrong");
                                }
                            }
                        }
                    }
                }
            }
            agg = null;
            if (recursive)
            {
                foreach (var folder in currentFolder.Folders.Cast<MAPIFolder>())
                {
                    await ProcessFolder(app, folder, recursive).ConfigureAwait(false);
                }
            }
    }
    }
}
