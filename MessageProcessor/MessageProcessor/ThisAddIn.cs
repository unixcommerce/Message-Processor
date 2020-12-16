using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace MessageProcessor
{
    public partial class ThisAddIn
    {
        public List<string> FilePathCollection = new List<string>();
        public int GlobalIndex = -1;
        public bool enableMessagg = false;
        Outlook.Application application;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            application = this.Application;
            application.ItemSend += Application_ItemSend;
        }

        private void Application_ItemSend(object Item, ref bool Cancel)
        {
            try
            {
                if (Item is Outlook.MailItem && enableMessagg)
                {
                    if (Globals.ThisAddIn.FilePathCollection.Count > Globals.ThisAddIn.GlobalIndex)
                        Globals.ThisAddIn.FilePathCollection.RemoveAt(Globals.ThisAddIn.GlobalIndex);
                    Globals.ThisAddIn.GlobalIndex = -1;
                    if (Globals.ThisAddIn.FilePathCollection.Count == 0)
                    {
                        DialogService.Services.displayDialogMessage("This is last message from Queue", DialogService.MessageType.Information);
                    }
                }
            }
            catch (Exception ex)
            {
                DialogService.Services.displayDialogMessage(ex.Message, DialogService.MessageType.Error);                
            }
        }

        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonItem();
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
