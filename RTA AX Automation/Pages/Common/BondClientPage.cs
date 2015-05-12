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
    public class BondClientPage
    {

        #region PageControls
        private WinWindow mUIAXCWindow;
        private WinClient mUIClientName;
        private static string windowName = "Bond clients";
        #endregion

        public class UIAXCWindow : WinWindow
        {
            public UIAXCWindow()
            {
                #region Search Criteria
                this.TechnologyName = "MSAA";
                this.SearchProperties.Add(new PropertyExpression(WinWindow.PropertyNames.Name, windowName, PropertyExpressionOperator.Contains));
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
                this.SearchProperties[WinWindow.PropertyNames.ControlId] = "12311";
                #endregion
            }
        }

     
        

        
        [ActionMethod]
        public void ClickCloseButton()
        {

            WinControl uIButtons = UIControls.GetControl("Close", "Button", new UIAXCWindow());
            Mouse.Click(uIButtons, new Point(uIButtons.Width / 2, uIButtons.Height / 2));

        }

        [ActionMethod]
        public void ClickSetup()
        {

            WinControl uIButtons = UIControls.GetControl("Setup", "Client", new UIAXCWindow());
            Mouse.Click(uIButtons, new Point(uIButtons.Width / 2, uIButtons.Height / 2));

        }

        [ActionMethod]
        public void ClickDeactivateMenuItem()
        {

            WinControl uIButtons = UIControls.GetDropDownControl("Deactivate", "MenuItem", new UIAXCWindow());
            Mouse.Click(uIButtons, new Point(uIButtons.Width / 2, uIButtons.Height / 2));

        }

            

        [ActionMethod]
        public bool GetWindowExistStatus()
        {
            this.mUIAXCWindow = new UIAXCWindow();
            WinClient uIClientName = new WinClient(mUIClientName);
            uIClientName.TechnologyName = "MSAA";
            uIClientName.SearchProperties.Add("ControlType", windowName);
            uIClientName.SearchProperties.Add("Name", "");
            mUIClientName = uIClientName;
            return true;
        }

        [ActionMethod]
        public WinTable GetClientOverviewTable()
        {
            return Table.GetTable("Grid", "12311", new UIAXCWindow());
        }

        [ActionMethod]
        public void ClickGeneralTab()
        {
            WinControl uIItem = UIControls.GetControl("General", "TabPage", new UIAXCWindow());
            Mouse.Click(uIItem, new Point(uIItem.Width / 2, uIItem.Height / 2));
        }

        private WinControl GetMethodOfPaymentControl()
        {
            WinControl uIItem = UIControls.GetControl("Method Of Payment", "Edit", new UIAXCWindow());
            return uIItem;
        }

        [ActionMethod]
        public bool IsMethodOfPaymentEditable()
        {
            WinControl uIItem = this.GetMethodOfPaymentControl();
            if (uIItem.GetProperty("ReadOnly").ToString().Contains("True"))
            {
                return true;
            }
            return false;
        }
    }
}
