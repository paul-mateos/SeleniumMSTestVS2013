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
    public class Homepage 
    {

        public Homepage()
        {
            WinWindow thisWindow = new UIAXCWindow();
            UIControls.ClickMaximizeButton(new UIAXCWindow());
        }


        #region PageControls
        private WinWindow mUIAXCWindow;
       
         #endregion

        public class UIAXCWindow : WinWindow
        {
            public UIAXCWindow()
            {
                #region Search Criteria
                this.TechnologyName = "MSAA";
                this.SearchProperties.Add(new PropertyExpression(WinWindow.PropertyNames.Name, "‪Microsoft Dynamics AX", PropertyExpressionOperator.Contains));
                this.SearchProperties.Add("ClassName", "AxMainFrame");
                this.SearchProperties.Add("ControlType", "Window");
                this.SearchProperties.Add("ControlId", "0");
                #endregion
                this.WaitForControlReady();
            }

        }

       

        [ActionMethod]
        public void ClickHomeTab()
        {

            if (this.GetTabExists("Home") == true)
            {
                WinControl uITabPage = UIControls.GetItemWindowControl("Home", "TabPage", new UIAXCWindow());
                uITabPage.WaitForControlReady();
                Mouse.Click(uITabPage, new Point(uITabPage.Width / 2, uITabPage.Height / 2));
            }else
            {
                throw new Exception("Tab does not exist");
            }
        }

        [ActionMethod]
        public bool GetTabExists(string tabName)
        {
            try
            {
                WinWindow parentWindow = new UIAXCWindow();
                WinWindow itemWindow = new WinWindow(parentWindow);
                itemWindow.SearchProperties.Add("Instance", "11");
                itemWindow.SearchProperties.Add("ClassName", "WindowsForms10.Window", PropertyExpressionOperator.Contains);
                itemWindow.WaitForControlReady();
                WinTabList tabList = new WinTabList(itemWindow);
                tabList.SearchProperties.Add("ControlType", "TabList");
                tabList.WaitForControlReady();
                UITestControlCollection tablistChildren = tabList.GetChildren();
                int tabListCount = tablistChildren.Count();
                for (int i = 0; i <= tabListCount; i++)
                {
                   
                    if ((tablistChildren.ElementAt(i).ControlType.ToString() == "TabPage") && (tablistChildren.ElementAt(i).Name == tabName) && !(tablistChildren.ElementAt(i).State.ToString().Contains("Invisible")))
                    {
                        return true;
                    }

                } return false;
            }
            catch
            {
                return false;
            }
           
        }


        [ActionMethod]
        public void ClickHomeTreeMenu()
        {

            WinControl uITabPage = UIControls.GetTreeItemControl("Home", new UIAXCWindow());
            uITabPage.WaitForControlReady();
            Mouse.Click(uITabPage, new Point(uITabPage.Width / 2, uITabPage.Height / 2));
        }

        [ActionMethod]
        public void ClickGeneralLedgerTab()
        {
           
            if (this.GetTabExists("General ledger") == true)
            {
                WinControl uITabPage = UIControls.GetItemWindowControl("General ledger", "TabPage", new UIAXCWindow());
                uITabPage.WaitForControlReady();
                Mouse.Click(uITabPage, new Point(uITabPage.Width / 2, uITabPage.Height / 2));
            }
            else
            {
                this.ClickConfigButton();
                WinControl uIMenuItem = UIControls.GetDropDownControl("General ledger", "MenuItem", new UIAXCWindow());
                Mouse.Click(uIMenuItem, new Point(uIMenuItem.Width / 2, uIMenuItem.Height / 2));
            }

            //Minimise all group controls

            UIControls.ClickMinimizeMenuGroup("Common", new UIAXCWindow());
            UIControls.ClickMinimizeMenuGroup("Journals", new UIAXCWindow());
            UIControls.ClickMinimizeMenuGroup("Inquiries", new UIAXCWindow());
            UIControls.ClickMinimizeMenuGroup("Reports", new UIAXCWindow());
            UIControls.ClickMinimizeMenuGroup("Periodic", new UIAXCWindow());
            UIControls.ClickMinimizeMenuGroup("Setup", new UIAXCWindow());
            

        }

        [ActionMethod]
        public void ClickAccountPayableTab()
        {
            if (this.GetTabExists("Accounts payable") == true)
            {
                WinControl uITabPage = UIControls.GetItemWindowControl("Accounts payable", "TabPage", new UIAXCWindow());
                uITabPage.WaitForControlReady();
                Mouse.Click(uITabPage, new Point(uITabPage.Width / 2, uITabPage.Height / 2));
            }
            else
            {
                this.ClickConfigButton();
                WinControl uIMenuItem = UIControls.GetDropDownControl("Accounts payable", "MenuItem", new UIAXCWindow());
                Mouse.Click(uIMenuItem, new Point(uIMenuItem.Width / 2, uIMenuItem.Height / 2));
            }
        }

        
        [ActionMethod]
        public void ClickConfigButton()
        {

            WinControl uIButtons = UIControls.GetItemWindowControl("Configure buttons", "Button", new UIAXCWindow());
            Mouse.Click(uIButtons, new Point(uIButtons.Width / 2, uIButtons.Height / 2));

        }

        [ActionMethod]
        public void ClickBondManagementTab()
        {
            if (this.GetTabExists("Bond management") == true)
            {
                WinControl uITabPage = UIControls.GetItemWindowControl("Bond management", "TabPage", new UIAXCWindow());
                uITabPage.WaitForControlReady();
                Mouse.Click(uITabPage, new Point(uITabPage.Width / 2, uITabPage.Height / 2));
            }
            else
            {
                this.ClickConfigButton();
                WinControl uIMenuItem = UIControls.GetDropDownControl("Bond management", "MenuItem", new UIAXCWindow());
                Mouse.Click(uIMenuItem, new Point(uIMenuItem.Width / 2, uIMenuItem.Height / 2));
            }

        }

        [ActionMethod]
        public void ClickBondManagementMenuItem()
        {
            if (this.GetTabExists("Bond management") == true)
            {
                WinControl uITabPage = UIControls.GetItemWindowControl("Bond management", "TabPage", new UIAXCWindow());
                uITabPage.WaitForControlReady();
                Mouse.Click(uITabPage, new Point(uITabPage.Width / 2, uITabPage.Height / 2));
            }
            else
            {
                WinControl uIMenuItem = UIControls.GetDropDownControl("Bond management", "MenuItem", new UIAXCWindow());
                Mouse.Click(uIMenuItem, new Point(uIMenuItem.Width / 2, uIMenuItem.Height / 2));
            }

        }

        [ActionMethod]
        public void ClickBondManagementMenuTreeItem()
        {
            WinTreeItem uITreeItem = new WinTreeItem(UIControls.GetTreeItemControl("Bond management", new UIAXCWindow()));
            if (uITreeItem.Expanded == false)
            {
                Mouse.Click(uITreeItem, new Point(uITreeItem.Width / 2, uITreeItem.Height / 2));
            }

        }

        [ActionMethod]
        public void ClickInquiriesMenuTreeItem()
        {
            WinTreeItem uITreeItem = new WinTreeItem(UIControls.GetTreeItemControl("Bond management", new UIAXCWindow()));
            if (uITreeItem.Expanded == false)
            {
                Mouse.Click(uITreeItem, new Point(25, uITreeItem.Height / 2));
            }

        }

        [ActionMethod]
        public void ClickCashandBankManagementMenuItem()
        {
            WinControl uIMenuItem = UIControls.GetDropDownControl("Cash and bank management", "MenuItem", new UIAXCWindow());
            Mouse.Click(uIMenuItem, new Point(uIMenuItem.Width / 2, uIMenuItem.Height / 2));

        }

        [ActionMethod]
        public void ClickCashandBankManagementTab()
        {

            if (this.GetTabExists("Cash and bank management") == true)
            {
                WinControl uITabPage = UIControls.GetItemWindowControl("Cash and bank management", "TabPage", new UIAXCWindow());
                uITabPage.WaitForControlReady();
                Mouse.Click(uITabPage, new Point(uITabPage.Width / 2, uITabPage.Height / 2));
            }
            else
            {
                this.ClickConfigButton();
                WinControl uIMenuItem = UIControls.GetDropDownControl("Cash and bank management", "MenuItem", new UIAXCWindow());
                Mouse.Click(uIMenuItem, new Point(uIMenuItem.Width / 2, uIMenuItem.Height / 2));
            }



        }

        [ActionMethod]
        public void ClickTrialBalanceMenuItem()
        {
            WinMenuItem uIMenuItem = this.GetMenuBarMenuItem("Trial balance", "MenuItem");
            Mouse.Click(uIMenuItem, new Point(uIMenuItem.Width / 2, uIMenuItem.Height / 2));

        }

        [ActionMethod]
        public void ClickSystemLink()
        {
            WinControl uILink = UIControls.GetGroupedControl("Submenu1", "Hyperlink", "Inquiries", new UIAXCWindow());
            if (uILink.Height <= 20)
            {
                Mouse.Click(uILink, new Point(25, uILink.Height / 2));
            }
        }

        [ActionMethod]
        public void ClickOutboundCRMMessagesLink()
        {
            WinWindow parentWindow = new UIAXCWindow();
            WinControl parentLink = UIControls.GetGroupedControl("Submenu1", "Hyperlink", "Inquiries", parentWindow);
            WinControl uILink = UIControls.GetChildControl("CRMOutboundNotification", parentLink, "Hyperlink");
            Mouse.Click(uILink, new Point(25, uILink.Height / 2));
            
        }
        

        [ActionMethod]
        public void ClickPaymentsLink()
        {
            WinControl uILink = UIControls.GetGroupedControl("BPAY", "Hyperlink", "Inquiries", new UIAXCWindow());
            if (uILink.Height <= 20)
            {
                Mouse.Click(uILink, new Point(25, uILink.Height / 2));
            }
        }

      
        [ActionMethod]
        public void ClickJournalsLink()
        {
            WinControl uILink = UIControls.GetGroupedControl("Journals", "Hyperlink", "Setup", new UIAXCWindow());
            if (uILink.Height <= 20)
            {
                Mouse.Click(uILink, new Point(25, uILink.Height / 2));
            }
        }

        

        [ActionMethod]
        public void ClickBankStatementFileImportExceptionLink()
        {
            WinControl uILink = UIControls.GetGroupedControl("BankStatementFileException", "Hyperlink", "Inquiries", new UIAXCWindow());
            Mouse.Click(uILink, new Point(25, uILink.Height / 2));
            
        }

        

        [ActionMethod]
        public void ClickConsolidateLink()
        {
            WinControl uILink = UIControls.GetGroupedControl("Consolidate", "Hyperlink", "Periodic", new UIAXCWindow());
            if (uILink.Height <= 20)
            {
                Mouse.Click(uILink, new Point(30, uILink.Height / 2));
            }
        }

        [ActionMethod]
        public void ClickGeneralJournalLink()
        {
            WinControl uILink = UIControls.GetGroupedControl("LedgerJournalTable3", "Hyperlink", "Journals", new UIAXCWindow());
            Mouse.Click(uILink, new Point(25, uILink.Height / 2));
           
        }

        [ActionMethod]
        public void ClickLedgerPostingGroupsLink()
        {

            WinWindow parentWindow = new UIAXCWindow();
            WinControl parentLink = UIControls.GetGroupedControl("SalesTax", "Hyperlink", "Setup", parentWindow);
            WinControl uILink = UIControls.GetChildControl("TaxAccountGroup", parentLink, "Hyperlink");
            Mouse.Click(uILink, new Point(25, uILink.Height / 2));

        }

        [ActionMethod]
        public void ClickImportEFTFileLink()
        {
            WinControl uILink = UIControls.GetControl("BondEFTFileImport", "Hyperlink", new UIAXCWindow());
            Mouse.Click(uILink, new Point(25, uILink.Height / 2));
        }

        [ActionMethod]
        public void ClickSalesTaxLink()
        {
            WinWindow parent = new UIAXCWindow();
            UIControls.ClickMaximizeMenuGroup("Setup", parent);
            WinControl uILink = UIControls.GetGroupedControl("SalesTax", "Hyperlink", "Setup", parent);
            if (uILink.Height <= 20)
            {
                Mouse.Click(uILink, new Point(25, uILink.Height / 2));
            }
        }

        [ActionMethod]
        public void ClickVendorsLink()
        {
            WinControl uILink = UIControls.GetGroupedControl("Vendors", "Hyperlink", "Common", new UIAXCWindow());
            if (uILink.Height <= 20)
            {
                Mouse.Click(uILink, new Point(25, uILink.Height / 2));
            }
        }

        [ActionMethod]
        public void ClickVendorsSetupLink()
        {
            WinControl uILink = UIControls.GetGroupedControl("Vendors", "Hyperlink", "Setup", new UIAXCWindow());
            if (uILink.Height <= 20)
            {
                Mouse.Click(uILink, new Point(25, uILink.Height / 2));
            }
        }

        [ActionMethod]
        public void ClickImportStatementButton()
        {
            WinControl uIItem = UIControls.GetTabGroupedControl("Import statement", "MenuItem", "HomeTab", "Import", new UIAXCWindow());
            Mouse.Click(uIItem, new Point(uIItem.Width / 2, uIItem.Height / 2));
        }

        [ActionMethod]
        public void ClickBankStatementsLink()
        {
            WinControl uILink = UIControls.GetGroupedControl("BankStatementTableListPage", "Hyperlink", "Common", new UIAXCWindow());
            Mouse.Click(uILink, new Point(25, uILink.Height / 2));
        }

        [ActionMethod]
        public void ClickBondClientLink()
        {
            WinControl uILink = UIControls.GetGroupedControl("BondClient", "Hyperlink", "Common", new UIAXCWindow());
            Mouse.Click(uILink, new Point(25, uILink.Height / 2));
        }

        [ActionMethod]
        public void ClickTrialBalanceLink()
        {
            WinControl uILink = UIControls.GetControl("LedgerTrialBalanceListPage", "Hyperlink", new UIAXCWindow());
            Mouse.Click(uILink, new Point(25, uILink.Height / 2));
        }
        

        [ActionMethod]
        public void ClickReceiptJournalsLink()
        {
            WinControl uILink = UIControls.GetGroupedControl("BondJournalTableReceipt", "Hyperlink", "Journal", new UIAXCWindow());
            Mouse.Click(uILink, new Point(25, uILink.Height / 2));
        }
        

        [ActionMethod]
        public void ClickBPayFileLink()
        {
            WinWindow parentWindow = new UIAXCWindow();
            WinControl parentLink = UIControls.GetGroupedControl("BPAY", "Hyperlink", "Inquiries", parentWindow);
            WinControl uILink = UIControls.GetChildControl("BPAYFileHeader", parentLink, "Hyperlink");
            Mouse.Click(uILink, new Point(25, uILink.Height / 2));
            
        }

        [ActionMethod]
        public void ClickConsolidateOnlineLink()
        {
            WinWindow parentWindow = new UIAXCWindow();
            WinControl parentLink = UIControls.GetGroupedControl("Consolidate", "Hyperlink", "Periodic", parentWindow);
            WinControl uILink = UIControls.GetChildControl("LedgerConsolidate2", parentLink, "Hyperlink");
            Mouse.Click(uILink, new Point(25, uILink.Height / 2)); 
           
        }

        
        [ActionMethod]
        public void ClickBPayReversalLink()
        {
            WinControl uILink = UIControls.GetControl("BPAYReversalPayment", "Hyperlink", new UIAXCWindow());
            Mouse.Click(uILink, new Point(25, uILink.Height / 2));
        }
        
        [ActionMethod]
        public void ClickBPayFileExceptionLink()
        {
            WinControl uILink = UIControls.GetControl("BPAYFileLineException", "Hyperlink", new UIAXCWindow());
            Mouse.Click(uILink, new Point(25, uILink.Height / 2));
        }

        [ActionMethod]
        public void ClickCRMOutboundNotificationsLink()
        {
            WinControl uILink = UIControls.GetGroupedControl("CRMOutbound", "Hyperlink", "Periodic", new UIAXCWindow());
            Mouse.Click(uILink, new Point(25, uILink.Height / 2));
        }

        [ActionMethod]
        public void ClickReceiptJournalLink()
        {
            WinControl uILink = UIControls.GetControl("BondJournalTableReceipt", "Hyperlink", new UIAXCWindow());
            Mouse.Click(uILink, new Point(25, uILink.Height / 2));
        }

        [ActionMethod]
        public void ClickJournalNamesLink()
        {
            WinControl uILink = UIControls.GetControl("LedgerJournalSetup", "Hyperlink", new UIAXCWindow());
            Mouse.Click(uILink, new Point(25, uILink.Height / 2));
        }

  

        [ActionMethod]
        public void ClickCompanyButton()
        {

            WpfButton uIButton = new WpfButton(UIControls.GetUIAControl("CompanyButton", "Button", new UIAXCWindow()));
            Mouse.Click(uIButton, new Point(uIButton.Width / 2, uIButton.Height / 2));
            
        }


     

        [ActionMethod]
        public WinMenuItem GetMenuBarMenuItem(string name, string type)
        {

            this.mUIAXCWindow = new UIAXCWindow();
            WinMenuBar mUIMenuBar = new WinMenuBar(mUIAXCWindow);
            mUIMenuBar.WaitForControlReady();
            WinMenuItem uIControl = new WinMenuItem(mUIMenuBar);
            uIControl.SearchProperties.Add("ControlType", "MenuItem");
            uIControl.SearchProperties.Add("Name", "Trial balance");
            uIControl.WaitForControlReady();
            return uIControl;
           

        }


        internal void ClickBackNavButton()
        {
            WinControl uIButton = UIControls.GetControl("Back", "Button", new UIAXCWindow());
            Mouse.Click(uIButton, new Point(25, uIButton.Height / 2));
            
        }
    }
}
