using System;
using Microsoft.Office.Interop.Outlook;
using System.Diagnostics;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace RemoveMailDuplicates
{
    internal class FlatFolder
    {
        private static async Task ProcessFolder(Application app, MAPIFolder mainFolder, MAPIFolder currentFolder, bool recursive)
        {
            Debug.Print(currentFolder.FullFolderPath);
            List<MailItem> list = new List<MailItem>(currentFolder.Items.Count);
            
            foreach (var item in currentFolder.Items)
            {
                var mail = item as MailItem;
                if (mail != null)
                {
                    list.Add(mail);
                }
                
            }

            foreach (var item in list)
            {
                item.Move(mainFolder);
            }

            if (recursive)
            {
                foreach (var folder in currentFolder.Folders.Cast<MAPIFolder>())
                {
                    await ProcessFolder(app, mainFolder, folder, recursive).ConfigureAwait(false);
                }
            }

        }

        internal static async void ProcessFolder(Application app, MAPIFolder currentFolder, bool recursive)
        {
            Debug.Print(currentFolder.FullFolderPath);
            foreach (var folder in currentFolder.Folders.Cast<MAPIFolder>())
            {
                await ProcessFolder(app, currentFolder, folder, recursive).ConfigureAwait(false);
            }

        }

    }
}
