using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace MessageProcessor.DialogService
{
    public class Services
    {
       public static void displayDialogMessage(string message, MessageType messageType)
        {
            MessageBoxIcon messageBoxIcon = MessageBoxIcon.Information;
            switch (messageType)
            {
                case MessageType.Error:
                    messageBoxIcon = MessageBoxIcon.Error;
                    break;
                case MessageType.Warning:
                    messageBoxIcon = MessageBoxIcon.Warning;
                    break;
                case MessageType.Information:
                    messageBoxIcon = MessageBoxIcon.Information;
                    break;
            }
            MessageBox.Show(message, "Message Processor", MessageBoxButtons.OK, messageBoxIcon);
        }
    }
    public enum MessageType
    {
        Error,
        Warning,
        Information
    }
}
