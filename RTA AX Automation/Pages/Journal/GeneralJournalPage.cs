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
    public class GeneralJournalPage
    {

        public GeneralJournalPage()
        {
            WinWindow thisWindow = new UIAXCWindow();
            UIControls.ClickMaximizeButton(new UIAXCWindow());
           
        }
    
        public class UIAXCWindow : WinWindow
        {
            public UIAXCWindow()
            {
                #region Search Criteria
                this.SearchProperties.Add("Name", "General journal‬ (‎‪1‬ - ‎‪rta‬)");//, PropertyExpressionOperator.Contains);
                this.SearchProperties.Add("ClassName", "AxTopLevelFrame");
                this.WindowTitles.Add("General journal‬ (‎‪1‬ - ‎‪rta‬)");
                #endregion

            }

        }

        
        public class UIItemWindow : WinWindow
        {

            public UIItemWindow(UITestControl searchLimitContainer) :
                base(searchLimitContainer)
            {
                #region Search Criteria
                this.SearchProperties[WinWindow.PropertyNames.ControlId] = "12336";
                #endregion
            }
        }
       

        [ActionMethod]
        public void SetNameValue(string value)
        {
            UIControls.SetItemControlValue("Name", "Edit", value, new UIAXCWindow());
            Keyboard.SendKeys("{TAB}");
        }

        [ActionMethod]
        public void SetDescriptionValue(string value)
        {
            UIControls.SetItemControlValue("Description", "Edit", value, new UIAXCWindow());
        }





        [ActionMethod]
        public void ClickLinesMenuItem()
        {

            WinControl uIMenuItem = UIControls.GetControl("Lines", "MenuItem", new UIAXCWindow());
            Mouse.Click(uIMenuItem, new Point(uIMenuItem.Width / 2, uIMenuItem.Height / 2));

        }

        [ActionMethod]
        public void ClickNewMenuItem()
        {

            WinControl uIMenuItem = UIControls.GetGroupedControl("New", "MenuItem", "NewDeleteGroup", new UIAXCWindow());
            Mouse.Click(uIMenuItem, new Point(uIMenuItem.Width / 2, uIMenuItem.Height / 2));

        }

      

        [ActionMethod]
        public void ClickCloseButton()
        {

            WinControl uIButton = UIControls.GetControl("Close", "Button", new UIAXCWindow());
            Mouse.Click(uIButton, new Point(uIButton.Width / 2, uIButton.Height / 2));

        }

      


        [ActionMethod]
        public WinTable GetBondReceiptTable()
        {
            return Table.GetTable("Grid", "12336", new UIAXCWindow());
        }


    }
}
