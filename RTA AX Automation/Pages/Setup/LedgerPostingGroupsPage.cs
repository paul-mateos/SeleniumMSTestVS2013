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
using System.Windows.Automation;
using RTA.Automation.AX.UI;


namespace RTA.Automation.AX.Pages
{
    [ActionClass]
    public class LedgerPostingGroupsPage
    {

        #region PageControls
        
        #endregion

        public class UIAXCWindow : WinWindow
        {
            public UIAXCWindow()
            {
                #region Search Criteria
                //this.TechnologyName = "MSAA";
                this.SearchProperties.Add("Name", "Ledger posting groups‬ (‎‪1‬ - ‎‪rta‬)‬‬", PropertyExpressionOperator.Contains);
                this.SearchProperties.Add("ClassName", "AxTopLevelFrame");
                //this.WindowTitles.Add("‪Ledger posting groups‬ (‎‪1‬ - ‎‪rta‬)");
                #endregion

            }

        }

        
        public class UIItemWindow : WinWindow
        {

            public UIItemWindow(UITestControl searchLimitContainer) :
                base(searchLimitContainer)
            {
                #region Search Criteria
                this.SearchProperties[WinWindow.PropertyNames.ControlId] = "12330";
                #endregion
            }
        }

       


        [ActionMethod]
        public void ClickLinesMenuItem()
        {

            WinControl uIMenuItem = UIControls.GetControl("Lines", "MenuItem", new UIAXCWindow());
            Mouse.Click(uIMenuItem, new Point(uIMenuItem.Width / 2, uIMenuItem.Height / 2));

        }

        [ActionMethod]
        public string GetSalesTaxPayableValue()
        {

            WinEdit uIItem = new WinEdit(UIControls.GetControl("Description", "Edit", new UIAXCWindow()));
            Mouse.Click(uIItem, new Point(uIItem.Width / 2, uIItem.Height / 2));
            Keyboard.SendKeys("{TAB}");

            WpfEdit uIAItem = new WpfEdit(UIControls.GetUIAControl("SegmentTextBox0", "Edit", new UIAXCWindow()));
            return uIAItem.Text;

        }

        [ActionMethod]
        public string GetSalesTaxReceivableValue()
        {

            WinEdit uIItem = new WinEdit(UIControls.GetControl("Description", "Edit", new UIAXCWindow()));
            Mouse.Click(uIItem, new Point(uIItem.Width / 2, uIItem.Height / 2));
            Keyboard.SendKeys("{TAB}");
            Keyboard.SendKeys("{TAB}");
            WpfText uIAText = new WpfText(UIControls.GetUIAControl("StatusTextBlock", "Text", new UIAXCWindow()));

            if (uIAText.DisplayText == "Account number for sales tax receivable")
            {
                var window = AutomationElement.RootElement.FindFirst(TreeScope.Descendants, new System.Windows.Automation.PropertyCondition(AutomationElement.NameProperty, "‪Ledger posting groups‬", PropertyConditionFlags.IgnoreCase));


                WpfEdit uIAItem = new WpfEdit(UIControls.GetUIAControl("SegmentTextBox0", "Edit", new UIAXCWindow()));
                return uIAItem.Text;
            }
            else
            {
                throw new Exception("Incorrect field selected");
            }
            

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


    }
}
