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
    public class EFTFileImportPage
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
                this.SearchProperties.Add("Name", "Microsoft Dynamics AX (1)", PropertyExpressionOperator.Contains);
                this.SearchProperties.Add("ClassName", "AxTopLevelFrame");
                #endregion

            }

        }


        [ActionMethod]
        public bool GetWindowExistStatus()
        {
            this.mUIAXCWindow = new UIAXCWindow();
            WinClient uIClientName = new WinClient(mUIClientName);
            uIClientName.TechnologyName = "MSAA";
            uIClientName.SearchProperties.Add("ControlType", "Client");
            uIClientName.SearchProperties.Add("Name", "Import EFT file");
            mUIClientName = uIClientName; 
            return uIClientName.WaitForControlExist();
            
        }

        [ActionMethod]
        public void SetMoveFileCheckBox(bool value)
        {
            UIControls.SetControlValue("Move file after import", "CheckBox", value.ToString(), new UIAXCWindow());
            
        }
        [ActionMethod]
        public void SetProcessFileCheckBox(bool value)
        {
            UIControls.SetControlValue("Process file after import", "CheckBox", value.ToString(), new UIAXCWindow());

        }
        [ActionMethod]
        public void SetImportPathText(string value)
        {
            UIControls.SetControlValue("Import path", "Edit", value, new UIAXCWindow());

        }


        [ActionMethod]
        public void ClickOKButton()
        {

            WinControl uIButtons = UIControls.GetControl("OK", "Button", new UIAXCWindow());
            Mouse.Click(uIButtons, new Point(uIButtons.Width / 2, uIButtons.Height / 2));

        }
        [ActionMethod]
        public void SetMoveFileAfterImportCheckBox(string value)
        {

            UIControls.SetControlValue("Move file after import", "CheckBox", value, new UIAXCWindow());

        }
        [ActionMethod]
        public void SetProcessFileAfterImportCheckBox(string value)
        {

            UIControls.SetControlValue("Process file after import", "CheckBox", value, new UIAXCWindow());

        }

        [ActionMethod]
        public void SetImportPathEdit(string value)
        {

            UIControls.SetControlValue("Import path", "Edit", value, new UIAXCWindow());

        }


       
    }
}
