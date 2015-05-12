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
    public class VoucherTransactionsPage
    {

        public VoucherTransactionsPage()
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
                this.SearchProperties.Add("Name", "Voucher transactions‬", PropertyExpressionOperator.Contains);
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
                this.SearchProperties[WinWindow.PropertyNames.ControlId] = "12331";
                #endregion
            }
        }


        [ActionMethod]
        public void ClickCloseButton()
        {

            WinControl uIButton = UIControls.GetControl("Close", "Button", new UIAXCWindow());
            Mouse.Click(uIButton, new Point(uIButton.Width / 2, uIButton.Height / 2));

        }

        [ActionMethod]
        public string GetDescriptionText()
        {
            WinEdit uIEditItem = new WinEdit(UIControls.GetItemControl("Description", "Edit", new UIAXCWindow()));
            return uIEditItem.Text;

        }

        [ActionMethod]
        public string GetAccountNameText()
        {
            WinEdit uIEditItem = new WinEdit(UIControls.GetItemControl("Account name", "Edit", new UIAXCWindow()));
            return uIEditItem.Text;

        }




        [ActionMethod]
        public WinTable GetVoucherTransactionTable()
        {
            return Table.GetTable("OverviewGrid", "12331", new UIAXCWindow());
        }

    }
}
