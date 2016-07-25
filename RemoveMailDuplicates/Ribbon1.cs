using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Tools.Ribbon;
using System.Threading.Tasks;

namespace RemoveMailDuplicates
{
    public partial class Ribbon1
    {
        private void Ribbon1_Load(object sender, RibbonUIEventArgs e)
        {

        }

        private void btnRemoveDuplicates_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            if (app == null )                return;
            
            var ActiveExplorer = app.ActiveExplorer();
            if (ActiveExplorer == null)
            {
                return;
            }

            Task.Run(delegate { RemoveDuplicate.ProcessFolder(app, ActiveExplorer.CurrentFolder, true); });
        }

        private void btnFlat_Click(object sender, RibbonControlEventArgs e)
        {
            var app = Globals.ThisAddIn.Application;
            if (app == null) return;

            var ActiveExplorer = app.ActiveExplorer();
            if (ActiveExplorer == null)
            {
                return;
            }

            Task.Run(delegate { FlatFolder.ProcessFolder(app, ActiveExplorer.CurrentFolder, true); } ); 
        }
    }
}
