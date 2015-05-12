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
    public class BondTransactionsPage
    {
        public BondTransactionsPage()
        {
            WinWindow thisWindow = new UIAXCWindow();
            UIControls.ClickMaximizeButton(new UIAXCWindow());
        }

       public class UIAXCWindow : WinWindow
        {
            public UIAXCWindow()
            {
                #region Search Criteria
                this.TechnologyName = "MSAA";
                this.SearchProperties.Add("Name", "Bond transactions‬", PropertyExpressionOperator.Contains);
                this.SearchProperties.Add("ClassName", "AxTopLevelFrame");
                #endregion

            }

        }

        public class UIItemWindow : WinWindow
        {

            public UIItemWindow(UITestControl searchLimitContainer) :
                base(searchLimitContainer)
            {
                #region Search Criteria
                this.SearchProperties[WinWindow.PropertyNames.ControlId] = "12299";
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
        public WinTable GetBondTransactionTable()
        {
            return  Table.GetTable("Grid", "12299", new UIAXCWindow());
            
        }
        

    }
}
