using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;
using Word = Microsoft.Office.Interop.Word;

namespace MessageProcessor
{
    [ComVisible(true)]
    public class RibbonItem : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public RibbonItem()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            if (ribbonID == "Microsoft.Outlook.Explorer" || ribbonID == "Microsoft.Outlook.Mail.Compose")
                return GetResourceText("MessageProcessor.RibbonItem.xml");
            else
                return null;
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit https://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }
        public Bitmap GetImageIcon(Office.IRibbonControl control)
        {
            if (control.Id.Equals("attachedMessage"))
            {
                return Properties.Resources.uploadMessage;
            }
            else if (control.Id.Equals("uploadMessage"))
            {
                return Properties.Resources.pullMessage;
            }
            else
            {
                return null;
            }
        }
        public void On_ActionAddMessages(Office.IRibbonControl control)
        {
            FolderBrowserDialog folderBrowserDialog = new FolderBrowserDialog();
            DialogResult result = folderBrowserDialog.ShowDialog();
            if (result == DialogResult.OK)
            {
                try
                {
                    string[] filesCollection = Directory.GetFiles(folderBrowserDialog.SelectedPath);
                    if (filesCollection != null && filesCollection.Length > 0)
                    {
                        Globals.ThisAddIn.FilePathCollection.AddRange(filesCollection);
                        Globals.ThisAddIn.enableMessagg = true;
                        DialogService.Services.displayDialogMessage($"{filesCollection.Length} messages are added to Queue", DialogService.MessageType.Information);
                    }
                    else
                    {
                        DialogService.Services.displayDialogMessage("There is not any messsage template file found.Please choose another folder", DialogService.MessageType.Warning);
                    }
                }
                catch (Exception ex)
                {
                    DialogService.Services.displayDialogMessage(ex.Message, DialogService.MessageType.Error);
                }
            }
        }
        public void On_Action_PullMessage(Office.IRibbonControl control)
        {
            try
            {
                var inspector = Globals.ThisAddIn.Application.ActiveInspector();
                if (inspector != null)
                {
                    var item = inspector.CurrentItem;
                    if (item is Outlook.MailItem)
                    {
                        var mailitem = item as Outlook.MailItem;
                        Random rnd = new Random();
                        if (Globals.ThisAddIn.FilePathCollection.Count > 0)
                        {
                            Globals.ThisAddIn.GlobalIndex = rnd.Next(0, Globals.ThisAddIn.FilePathCollection.Count);
                        }
                        if (Globals.ThisAddIn.GlobalIndex > -1)
                        {
                            try
                            {
                                string file = Globals.ThisAddIn.FilePathCollection[Globals.ThisAddIn.GlobalIndex];
                                string text = File.ReadAllText(file);
                                Word.Document document = (Word.Document)mailitem.GetInspector.WordEditor;
                                int postion = document.Application.Selection.Range.Start;
                                document.Application.Selection.Range.InsertAfter(text);                                
                                Globals.ThisAddIn.GlobalIndex = -1;
                            }
                            catch (Exception ex)
                            {
                                DialogService.Services.displayDialogMessage(ex.Message, DialogService.MessageType.Error);
                            }
                        }
                        else
                        {                            
                            DialogService.Services.displayDialogMessage("Please add message in Queue from Home Tab using Add Message button", DialogService.MessageType.Information);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                DialogService.Services.displayDialogMessage(ex.Message, DialogService.MessageType.Error);
            }
        }
        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        #endregion
    }
}
