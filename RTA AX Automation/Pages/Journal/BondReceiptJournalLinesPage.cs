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
    public class BondReceiptJournalLinesPage
    {

        public BondReceiptJournalLinesPage()
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
                this.SearchProperties.Add("Name", "Bond receipt journal lines‬ (‎‪1‬ - ‎‪rtb‬)‎‪ - ‎‪Journal type: ReceiptBond journal:", PropertyExpressionOperator.Contains);
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
                this.SearchProperties[WinWindow.PropertyNames.ControlId] = "12307";
                #endregion
            }
        }














        [ActionMethod]
        public void ClickBondTransactionsMenuItem()
        {

            WinControl uIMenuItem = UIControls.GetDropDownControl("Bond transactions", "MenuItem", new UIAXCWindow());
            Mouse.Click(uIMenuItem, new Point(uIMenuItem.Width / 2, uIMenuItem.Height / 2));
            
        }

        [ActionMethod]
        public void ClickVoucherTransactionsMenuItem()
        {

            WinControl uIMenuItem =UIControls.GetDropDownControl("Voucher transactions", "MenuItem", new UIAXCWindow());
            Mouse.Click(uIMenuItem, new Point(uIMenuItem.Width / 2, uIMenuItem.Height / 2));

        }

        [ActionMethod]
        public void ClickInquiryButton()
        {

            WinControl uIControl = UIControls.GetControl("Inquiry", "MenuButton", new UIAXCWindow());
            Mouse.Click(uIControl, new Point(uIControl.Width / 2, uIControl.Height / 2));

        }

        [ActionMethod]
        public void ClickCloseButton()
        {

            WinControl uIButton = UIControls.GetControl("Close", "Button", new UIAXCWindow());
            Mouse.Click(uIButton, new Point(uIButton.Width / 2, uIButton.Height / 2));

        }

        [ActionMethod]
        public void SetShowAllText(string value)
        {

            UIControls.SetControlValue("Show", "Edit", value, new UIAXCWindow());

        }




        [ActionMethod]
        public WinTable GetBondReceiptJournalLinesTable()
        {

            return Table.GetTable("Grid", "12307", new UIAXCWindow());
           
        }



    }
}
