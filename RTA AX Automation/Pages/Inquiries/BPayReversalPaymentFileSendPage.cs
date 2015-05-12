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
    public class BPayReversalPaymentFileSendPage
    {

     #region PageControls
        private WinWindow mUIAXCWindow;
        private UITestControl mUIClientName;
    
        #endregion

        public class UIAXCWindow : WinWindow
        {
            public UIAXCWindow()
            {
                #region Search Criteria
                this.SearchProperties[WinWindow.PropertyNames.Name] = "‪Microsoft Dynamics AX‬ (‎‪1‬)‎";
                this.SearchProperties[WinWindow.PropertyNames.ClassName] = "AxTopLevelFrame";
                this.WindowTitles.Add("‪Microsoft Dynamics AX‬ (‎‪1‬)‎");
                #endregion
            }

        }

        public class UIItemWindow : WinWindow
        {

            public UIItemWindow(UITestControl searchLimitContainer) :
                base(searchLimitContainer)
            {
                #region Search Criteria
                this.SearchProperties[WinWindow.PropertyNames.ControlId] = "12297";
                this.WindowTitles.Add("‪Microsoft Dynamics AX‬ (‎‪1‬)‎");
                #endregion
            }
        }
   
        [ActionMethod]
        public bool GetWindowExistStatus()
        {
            this.mUIAXCWindow = new UIAXCWindow();
            WinClient uIClientName = new WinClient(mUIClientName);
            uIClientName.TechnologyName = "MSAA";
            uIClientName.SearchProperties.Add("ControlType", "BPAY Reversal Payment file send");
            uIClientName.SearchProperties.Add("Name", "");
            mUIClientName = uIClientName;
            return true;
        }


        [ActionMethod]
        public void ClickOKButton()
        {

            WinControl uIButtons = UIControls.GetControl("OK", "Button", new UIAXCWindow());
            Mouse.Click(uIButtons, new Point(uIButtons.Width / 2, uIButtons.Height / 2));

        }

    }
}
