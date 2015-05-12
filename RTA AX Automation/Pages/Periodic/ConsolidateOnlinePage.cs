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
    public class ConsolidateOnlinePage
    {

        #region PageControls
        private WinWindow mUIAXCWindow;
       
        #endregion

        public class UIAXCWindow : WinWindow
        {
            public UIAXCWindow()
            {
                #region Search Criteria
                this.TechnologyName = "MSAA";
                this.SearchProperties.Add("Name", "Consolidate [Online]‬");//, PropertyExpressionOperator.Contains);
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
                this.SearchProperties[WinWindow.PropertyNames.ControlId] = "12303";
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
                #endregion
            }

        }
         

        [ActionMethod]
        public void ClickOKButton()
        {

            WinControl uIButton = UIControls.GetControl("OK", "Button", new UIAXCWindow());
            Mouse.Click(uIButton, new Point(uIButton.Width / 2, uIButton.Height / 2));

        }

        [ActionMethod]
        public void SetFromValue(string value)
        {
            UIControls.SetItemControlValue("From:", "Edit", value, new UIAXCWindow());
        }

        [ActionMethod]
        public void SetToValue(string value)
        {
            UIControls.SetItemControlValue("To:", "Edit", value, new UIAXCWindow());
        }

        [ActionMethod]
        public void ClickFinancialdimensionsTab()
        {
            WinControl uIItem = UIControls.GetControl("Financial dimensions", "TabPage", new UIAXCWindow());
            Mouse.Click(uIItem, new Point(uIItem.Width / 2, uIItem.Height / 2));
        }

        [ActionMethod]
        public void ClickLegalentitiesTab()
        {
            WinControl uIItem = UIControls.GetControl("Legal entities", "TabPage", new UIAXCWindow());
            Mouse.Click(uIItem, new Point(uIItem.Width / 2, uIItem.Height / 2));
        }

        [ActionMethod]
        public void ClickDescriptionTab()
        {
            WinControl uIItem = UIControls.GetControl("Description", "TabPage", new UIAXCWindow());
            Mouse.Click(uIItem, new Point(uIItem.Width / 2, uIItem.Height / 2));
        }

        [ActionMethod]
        public void ClickEliminationTab()
        {
            WinControl uIItem = UIControls.GetControl("Elimination", "TabPage", new UIAXCWindow());
            Mouse.Click(uIItem, new Point(uIItem.Width / 2, uIItem.Height / 2));
        }

        [ActionMethod]
        public void ClickCriteriaTab()
        {
            WinControl uIItem = UIControls.GetControl("Criteria", "TabPage", new UIAXCWindow());
            Mouse.Click(uIItem, new Point(uIItem.Width / 2, uIItem.Height / 2));
        }

        [ActionMethod]
        public void ClickYesConsolidatedMessageBoxButton()
        {
            mUIAXCWindow = new WinWindow();
            mUIAXCWindow.SearchProperties.Add("Name", "Microsoft Dynamics");
            mUIAXCWindow.SearchProperties.Add("ClassName", "#32770");
            mUIAXCWindow.WindowTitles.Add("Microsoft Dynamics");
            WinButton yesButton = new WinButton(mUIAXCWindow);
            yesButton.SearchProperties.Add("Name", "Yes");
            yesButton.SearchProperties.Add("ControlType", "Button");
            yesButton.WindowTitles.Add("Microsoft Dynamics");
            yesButton.WaitForControlReady();
            Mouse.Click(yesButton, new Point(yesButton.Width / 2, yesButton.Height / 2));
            
        }

       
   }
}
