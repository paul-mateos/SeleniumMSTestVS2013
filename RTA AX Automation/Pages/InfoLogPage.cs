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
    public class InfoLogPage
    {

        public InfoLogPage()
        {
            WinWindow thisWindow = new UIAXCWindow();
            UIControls.ClickMaximizeButton(new UIAXCWindow());
        }

        #region PageControls
        private WinWindow mUIAXCWindow;
        private WinTreeItem mUITreeItem;
        private WinClient mUIClientName;
        #endregion

        public class UIAXCWindow : WinWindow
        {
            public UIAXCWindow()
            {
                #region Search Criteria
                this.TechnologyName = "MSAA";
                this.SearchProperties.Add("Name", "Infolog‬", PropertyExpressionOperator.Contains);
                this.SearchProperties.Add("ClassName", "AxTopLevelFrame");
                #endregion

            }

        }

       

       

        [ActionMethod]
        public void ClickClearButton()
        {

            WinControl uIButtons = UIControls.GetControl("Clear", "Button", new UIAXCWindow());
            Mouse.Click(uIButtons, new Point(uIButtons.Width / 2, uIButtons.Height / 2));

        }

        [ActionMethod]
        public void ClickCloseButton()
        {

            WinControl uIButtons = UIControls.GetControl("Close", "Button", new UIAXCWindow());
            Mouse.Click(uIButtons, new Point(uIButtons.Width / 2, uIButtons.Height / 2));

        }

     

        [ActionMethod]
        public bool GetTreeItemNOTExists(string param, string parentTreeItem)
        {
            this.mUIAXCWindow = new UIAXCWindow();
            WinTreeItem uIParentTreeItem = new WinTreeItem(mUIAXCWindow);
            uIParentTreeItem.SearchProperties.Add("Name", parentTreeItem, PropertyExpressionOperator.Contains);
            uIParentTreeItem.WaitForControlReady();
            this.mUITreeItem = new WinTreeItem(uIParentTreeItem);
            mUITreeItem.SearchProperties.Add("Name", param, PropertyExpressionOperator.Contains);
            mUITreeItem.SearchConfigurations.Add(SearchConfiguration.ExpandWhileSearching);
            mUITreeItem.SearchConfigurations.Add(SearchConfiguration.NextSibling);
            if(mUITreeItem.WaitForControlNotExist())
            {
                return true;
            }
            else
            {
                return false;
            }
            
        }

    

        [ActionMethod]
        public bool GetTreeItemExists(string param, string parentTreeItem)
        {
            this.mUIAXCWindow = new UIAXCWindow();
            WinTreeItem uIParentTreeItem = new WinTreeItem(mUIAXCWindow);
            uIParentTreeItem.SearchProperties.Add("Name", parentTreeItem, PropertyExpressionOperator.Contains);
            uIParentTreeItem.WaitForControlReady();
            this.mUITreeItem = new WinTreeItem(uIParentTreeItem);
            mUITreeItem.SearchProperties.Add("Name", param, PropertyExpressionOperator.Contains);
            mUITreeItem.SearchConfigurations.Add(SearchConfiguration.ExpandWhileSearching);
            mUITreeItem.SearchConfigurations.Add(SearchConfiguration.NextSibling);
            mUITreeItem.WaitForControlExist();
            return true;
        }

        [ActionMethod]
        public bool GetControlExists(string name, string type)
        {
            try
            {
                WinControl uIControl = UIControls.GetItemControl(name, type, new UIAXCWindow());
                return true;
            }
            catch
            {
                return false;
            }


        }

        [ActionMethod]
        public bool GetWindowExistStatus()
        {
            this.mUIAXCWindow = new UIAXCWindow();
            WinClient uIClientName = new WinClient(mUIClientName);
            uIClientName.TechnologyName = "MSAA";
            uIClientName.SearchProperties.Add("ControlType", "InfoLog");
            uIClientName.SearchProperties.Add("Name", "");
            mUIClientName = uIClientName;
            return true;
        }

        [ActionMethod]
        public string GetTreeItemName(string param, string parentTreeItem)
        {
            this.mUIAXCWindow = new UIAXCWindow();
            WinTreeItem uIParentTreeItem = new WinTreeItem(mUIAXCWindow);
            uIParentTreeItem.SearchProperties.Add("Name", parentTreeItem, PropertyExpressionOperator.Contains);
            uIParentTreeItem.WaitForControlReady();
            this.mUITreeItem = new WinTreeItem(uIParentTreeItem);
            mUITreeItem.SearchProperties.Add("Name", param, PropertyExpressionOperator.Contains);
            mUITreeItem.SearchConfigurations.Add(SearchConfiguration.ExpandWhileSearching);
            mUITreeItem.SearchConfigurations.Add(SearchConfiguration.NextSibling);
            mUITreeItem.WaitForControlExist();
            return mUITreeItem.Name;
        }


    }
}
