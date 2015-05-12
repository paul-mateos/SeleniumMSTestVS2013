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
    public class BondReceiptJournalPage
    {

        public BondReceiptJournalPage()
        {
            WinWindow thisWindow = new UIAXCWindow();
            UIControls.ClickMaximizeButton(new UIAXCWindow());
        }

        public class UIAXCWindow : WinWindow
        {
            public UIAXCWindow()
            {
                #region Search Criteria
                //this.TechnologyName = "MSAA";
                this.SearchProperties.Add("Name", "Bond receipt journal‬ (‎‪1‬ - ‎‪rtb‬)");//, PropertyExpressionOperator.Contains);
                this.SearchProperties.Add("ClassName", "AxTopLevelFrame");
                this.WindowTitles.Add("‪Bond receipt journal‬ (‎‪1‬ - ‎‪rtb‬)‎");
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

        public class UIItemWindow2 : WinWindow
        {

            public UIItemWindow2(UITestControl searchLimitContainer) :
                base(searchLimitContainer)
            {
                #region Search Criteria
                this.SearchProperties[WinWindow.PropertyNames.ControlId] = "12304";
                this.SearchProperties[WinWindow.PropertyNames.Instance] = "3";
                this.WindowTitles.Add("‪Bond receipt journal‬ (‎‪1‬ - ‎‪rtb‬)‎");
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
        public void ClickRemoveFilterMenuItem()
        {

             UIControls.ClickContextMenuItem("Remove Filter/Sort");

        }

        [ActionMethod]
        public void ClickFilterMenuItem()
        {

             UIControls.ClickContextMenuItem("Filter by field");

        }

        [ActionMethod]
        public void ClickCloseButton()
        {

            WinControl uIButton =UIControls.GetControl("Close", "Button", new UIAXCWindow());
            Mouse.Click(uIButton, new Point(uIButton.Width / 2, uIButton.Height / 2));

        }

        [ActionMethod]
        public void SetShowAllText(string value)
        {
            UIControls.SetControlValue("Show", "Edit", value, new UIAXCWindow());
        }


       

        [ActionMethod]
        public WinTable GetBondReceiptTable()
        {
            return Table.GetTable("Grid", "12307", new UIAXCWindow());
        }




    }
}
