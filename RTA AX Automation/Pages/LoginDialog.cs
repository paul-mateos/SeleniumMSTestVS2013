using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using Mouse = Microsoft.VisualStudio.TestTools.UITesting.Mouse;
using MouseButtons = System.Windows.Forms.MouseButtons;
using ActionWordsLib.Attributes;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UITesting.WpfControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Drawing;


namespace RTA.Automation.AX.Pages
{
    [ActionClass]
    public class LoginDialog
    {
        private WinTabPage mUITabPage;
        private WinWindow mUIAXCWindow;
        [ActionMethod]
        public void Login(string usernameParam, string passwordParam)
        {
            // To generate code for this test, select "Generate Code for Coded UI Test" from the shortcut menu and select one of the menu items.
            // Find Dynamics AX Window
 
            WinWindow dynamicsAXWindow = new WinWindow();
            dynamicsAXWindow.TechnologyName = "MSAA";
            dynamicsAXWindow.SearchProperties.Add("Name", "Microsoft Dynamics AX", PropertyExpressionOperator.Contains);
            dynamicsAXWindow.SearchProperties.Add("ClassName", "AxMainFrame");
            dynamicsAXWindow.WaitForControlExist();
            mUIAXCWindow = dynamicsAXWindow;
  
            WinTabPage uITabPage = new WinTabPage(mUIAXCWindow);
            uITabPage.TechnologyName = "MSAA";
            uITabPage.SearchProperties.Add("ControlType", "TabPage");
            uITabPage.SearchProperties.Add("Name", "Home");
            uITabPage.WaitForControlReady();
            mUITabPage = uITabPage;
            mUITabPage.WaitForControlReady();
            Mouse.Click(uITabPage, new Point(uITabPage.Width / 2, mUITabPage.Height / 2));
        }
    
        
    }
}
