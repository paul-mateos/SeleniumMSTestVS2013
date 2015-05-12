using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using Mouse = Microsoft.VisualStudio.TestTools.UITesting.Mouse;
using MouseButtons = System.Windows.Forms.MouseButtons;
using RTA.Automation.AX.Utils;
using ActionWordsLib.Attributes;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UITesting.WpfControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RTA.Automation.AX.Pages;
using System.Drawing;
using RTA.Automation.AX.UI;



namespace RTA.Automation.AX.Pages
{
    [ActionClass]
    public class OutboundCRMMessagePage
    {

        public OutboundCRMMessagePage()
        {
            WinWindow thisWindow = new UIAXCWindow();
            UIControls.ClickMaximizeButton(new UIAXCWindow());
        }

        public class UIAXCWindow : WinWindow
        {
            public UIAXCWindow()
            {
                #region Search Criteria
                this.SearchProperties.Add(new PropertyExpression(WinWindow.PropertyNames.Name, "‪CRM outbound message queue", PropertyExpressionOperator.Contains));
                this.SearchProperties[WinWindow.PropertyNames.ClassName] = "AxTopLevelFrame";
                this.WindowTitles.Add("‪CRM outbound message queue‬ (‎‪1‬ - ‎‪rtb‬)‎");
                #endregion

            }

        }

        public class UIItemWindow : WinWindow
        {

            public UIItemWindow(UITestControl searchLimitContainer) :
                base(searchLimitContainer)
            {
                #region Search Criteria
                this.SearchProperties[WinWindow.PropertyNames.ControlId] = "12295";
                this.WindowTitles.Add("‪CRM outbound message queue‬ (‎‪1‬ - ‎‪rtb‬)‎");
                #endregion
            }
        }


      
        
        [ActionMethod]
        public void ClickCloseButton()
        {

            WinButton uIButton = new WinButton(UIControls.GetControl("Close", "Button", new UIAXCWindow()));
            Mouse.Click(uIButton, new Point(uIButton.Width / 2, uIButton.Height / 2));

        }

        [ActionMethod]
        public void ClickFilterMenuItem()
        {

            UIControls.ClickContextMenuItem("Filter by field");


        }

       

        [ActionMethod]
        public WinTable GetNotificationTable()
        {
            return Table.GetTable("Grid", "12295", new UIAXCWindow());
        }

       

    }
}
