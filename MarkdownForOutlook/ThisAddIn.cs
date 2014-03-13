using System.Collections.Generic;
using System.Windows.Forms;
using Kiwi.Markdown;
using MouseKeyboardActivityMonitor;
using MouseKeyboardActivityMonitor.WinApi;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace MarkdownForOutlook
{
    public partial class ThisAddIn
    {
	    private Dictionary<Outlook.MailItem, string> m_markdownUndo = new Dictionary<Outlook.MailItem, string>();
		private KeyboardHookListener m_keyboardListener;
		private readonly MarkdownService m_markdownProvider = new MarkdownService(null);
 
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
	        m_keyboardListener = new KeyboardHookListener(new AppHooker()) { Enabled = true };
	        m_keyboardListener.KeyDown += KeyboardListener_KeyDown;
        }

	    private void KeyboardListener_KeyDown(object sender, KeyEventArgs keyEventArgs)
	    {
		    if (keyEventArgs.Alt && keyEventArgs.Control && keyEventArgs.KeyCode == Keys.M)
		    {
				var mailItem = Application.ActiveInspector().CurrentItem as Outlook.MailItem;
			    if (mailItem != null)
			    {
				    if (!m_markdownUndo.ContainsKey(mailItem))
					    m_markdownUndo.Add(mailItem,null);

					// TODO: 
					// * recognize and show warning whenever undo text will clobber changes made by the user
					// * don't transform user signature and previous messages in the thread
					// * cleanup old mail items (will hooking into send work?)
				    string bodyTransformed;
				    if (m_markdownUndo[mailItem] != null)
				    {
					    bodyTransformed = m_markdownUndo[mailItem];
					    m_markdownUndo[mailItem] = null;
				    }
				    else
				    {
					    bodyTransformed = m_markdownProvider.ToHtml(mailItem.Body);
					    m_markdownUndo[mailItem] = mailItem.HTMLBody;
				    }
				    
					mailItem.HTMLBody = bodyTransformed;
					keyEventArgs.Handled = true;   
			    }
		    }
	    }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
	        m_keyboardListener.Dispose();
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