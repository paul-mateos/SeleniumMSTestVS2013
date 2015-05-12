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
    public class OutboundCRMIntegrationPage
    {

        #region PageControls
        private WinWindow mUIAXCWindow;
        private WinClient mUIClientName;
        #endregion

        public class UIAXCWindow : WinWindow
        {
            public UIAXCWindow()
            {
                #region Search Criteria
                this.TechnologyName = "MSAA";
                this.SearchProperties.Add("Name", "Microsoft Dynamics AX", PropertyExpressionOperator.Contains);
                this.SearchProperties.Add("ClassName", "AxTopLevelFrame");
                #endregion

            }

        }

  


        [ActionMethod]
        public void ClickOKButton()
        {

            WinControl uIButtons = UIControls.GetControl("OK", "Button", new UIAXCWindow());
            Mouse.Click(uIButtons, new Point(uIButtons.Width / 2, uIButtons.Height / 2));

        }

        [ActionMethod]
        public void ClickCloseButton()
        {

            WinControl uIButtons = UIControls.GetControl("Close", "Button", new UIAXCWindow());
            Mouse.Click(uIButtons, new Point(uIButtons.Width / 2, uIButtons.Height / 2));

        }


        [ActionMethod]
        public bool GetWindowExistStatus()
        {
            this.mUIAXCWindow = new UIAXCWindow();
            WinClient uIClientName = new WinClient(mUIClientName);
            uIClientName.SearchProperties.Add("ControlType", "Client");
            uIClientName.SearchProperties.Add("Name", "Outbound CRM integration");
            mUIClientName = uIClientName;
            return true;
        }

    }
}
