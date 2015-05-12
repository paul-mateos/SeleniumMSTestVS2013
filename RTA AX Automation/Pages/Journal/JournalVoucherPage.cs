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
    public class JournalVoucherPage
    {

        public JournalVoucherPage()
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
                this.SearchProperties.Add(new PropertyExpression(WinWindow.PropertyNames.Name, "Journal voucher", PropertyExpressionOperator.Contains));
                this.SearchProperties.Add("ClassName", "AxTopLevelFrame");
                #endregion
                this.WaitForControlReady();
            }

        }

        public class UIItemWindow : WinWindow
        {

            public UIItemWindow(UITestControl searchLimitContainer) :
                base(searchLimitContainer)
            {
                #region Search Criteria
                this.SearchProperties[WinWindow.PropertyNames.ControlId] = "12417";
                
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
        public void ClickPostMenuButton()
        {

            WinControl uIMenuButton = UIControls.GetButtonGroupControl("Post", "MenuButton", new UIAXCWindow());
            Mouse.Click(uIMenuButton, new Point(uIMenuButton.Width / 2, uIMenuButton.Height / 2));

        }

        [ActionMethod]
        public void ClickInquiriesMenuButton()
        {

            WinControl uIMenuButton = UIControls.GetButtonGroupControl("Inquiries", "MenuButton", new UIAXCWindow());
            Mouse.Click(uIMenuButton, new Point(uIMenuButton.Width / 2, uIMenuButton.Height / 2));

        }

        [ActionMethod]
        public void ClickPostItemMenuItem()
        {

            WinControl uIMenuItem = UIControls.GetDropDownControl("Post", "MenuItem", new UIAXCWindow());
            Mouse.Click(uIMenuItem, new Point(uIMenuItem.Width / 2, uIMenuItem.Height / 2));

        }

        [ActionMethod]
        public void ClickVoucherItemMenuItem()
        {

            WinControl uIMenuItem =UIControls.GetDropDownControl("Voucher", "MenuItem", new UIAXCWindow());
            Mouse.Click(uIMenuItem, new Point(uIMenuItem.Width / 2, uIMenuItem.Height / 2));

        }


        [ActionMethod]
        public void ClickReportAsReadyItemMenuItem()
        {

            WinControl uIMenuItem = UIControls.GetDropDownControl("Report as ready", "MenuItem", new UIAXCWindow());
            Mouse.Click(uIMenuItem, new Point(uIMenuItem.Width / 2, uIMenuItem.Height / 2));

        }

        [ActionMethod]
        public void ClickApproveItemMenuItem()
        {

            WinControl uIMenuItem = UIControls.GetDropDownControl("Approve", "MenuItem", new UIAXCWindow());
            Mouse.Click(uIMenuItem, new Point(uIMenuItem.Width / 2, uIMenuItem.Height / 2));

        }
        

        [ActionMethod]
        public void ClickApprovalItemMenuButton()
        {

            WinControl uIMenuButton = UIControls.GetButtonGroupControl("Approval", "MenuButton", new UIAXCWindow());
            Mouse.Click(uIMenuButton, new Point(uIMenuButton.Width / 2, uIMenuButton.Height / 2));

        }

        [ActionMethod]
        public string GetDescriptionText()
        {
            return UIControls.GetControlValue("Description", "Edit", new UIAXCWindow());

        }

        [ActionMethod]
        public string GetAccountNameText()
        {
            return UIControls.GetControlValue("Account name", "Edit", new UIAXCWindow());

        }

        [ActionMethod]
        public string GetCalculatedSalesTaxAmountText()
        {
            WinEdit uIEdit = new WinEdit(UIControls.GetItemControl("Calculated GST amount", "Edit", new UIAXCWindow()));
            return uIEdit.Text;

        }

        [ActionMethod]
        public void SetAccountSeg1Value(string value)
        {
            UIControls.SetUIAItemControlValue("SegmentTextBox0", "Edit", value, new UIAXCWindow());
            //UIControls.SetUIAControlValue("SegmentTextBox0", "Edit", value, new UIAXCWindow());
            Keyboard.SendKeys("{TAB}");
        
        }

        [ActionMethod]
        public void SetAccountSeg2Value(string value)
        {
            UIControls.SetUIAItemControlValue("SegmentTextBox1", "Edit", value, new UIAXCWindow());
            Keyboard.SendKeys("{TAB}");

        }
        [ActionMethod]
        public void SetAccountSeg3Value(string value)
        {
            UIControls.SetUIAItemControlValue("SegmentTextBox2", "Edit", value, new UIAXCWindow());
            Keyboard.SendKeys("{TAB}");

        }


        [ActionMethod]
        public void SetDebitValue(string value)
        {
            UIControls.SetItemControlValue("Debit", "Edit", value,new UIAXCWindow());

        }

        [ActionMethod]
        public void SetSalesTaxGroupValue(string value)
        {
            UIControls.SetItemControlValue("GST group", "Edit", value, new UIAXCWindow());

        }

        [ActionMethod]
        public void SetItemGSTGroupValue(string value)
        {
            UIControls.SetItemControlValue("Item GST group", "Edit", value, new UIAXCWindow());

        }

        [ActionMethod]
        public WinTable GetJournalValueTable()
        {


            return Table.GetTable("overviewGrid", "12417", new UIAXCWindow());

        }

    }
}
