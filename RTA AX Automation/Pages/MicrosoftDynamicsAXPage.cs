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
    public class MicrosoftDynamicsAXPage
    {


        public class UIAXCWindow : WinWindow
        {
            public UIAXCWindow()
            {
                #region Search Criteria
                this.TechnologyName = "MSAA";
                this.SearchProperties.Add("Name", "Microsoft Dynamics AX (1)", PropertyExpressionOperator.EqualTo);
                this.SearchProperties.Add("ClassName", "AxTopLevelFrame");
                #endregion

            }

        }


        [ActionMethod]
        public void ClickOkButton()
        {

            WinControl uIButtons =UIControls.GetControl("Ok", "Button", new UIAXCWindow());
            Mouse.Click(uIButtons, new Point(uIButtons.Width / 2, uIButtons.Height / 2));

        }



        [ActionMethod]
        public void SetDeactivationReasonText(string text)
        {
            UIControls.SetControlValue("Deactivation reason", "Edit", text, new UIAXCWindow());

        }

       
 

    }
}
