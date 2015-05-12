using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Keyboard = Microsoft.VisualStudio.TestTools.UITesting.Keyboard;
using Mouse = Microsoft.VisualStudio.TestTools.UITesting.Mouse;
using MouseButtons = System.Windows.Forms.MouseButtons;
using RTA_AX_Automation.Utils;
using ActionWordsLib.Attributes;
using Microsoft.VisualStudio.TestTools.UITest.Extension;
using Microsoft.VisualStudio.TestTools.UITesting;
using Microsoft.VisualStudio.TestTools.UITesting.WinControls;
using Microsoft.VisualStudio.TestTools.UITesting.WpfControls;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using RTA_AX_Automation.UIMaps;
using RTA_AX_Automation.Pages;
using System.Drawing;



namespace RTA_AX_Automation.Pages
{
    [ActionClass]
    public class BPayReversalPaymentFileSend
    {
        [ActionMethod]
        public static bool GetWindowExistStatus(string windowTitle)
        {
            if (Instantiator.MSDynamicsAXUI.GetWindowExistStatus(windowTitle) == true)
            {
                return true;
            }else
            {
                return false;
            }

        }
        [ActionMethod]
        public static void ClickTab(string param)
        {
            Instantiator.MSDynamicsAXUI.ClickTabPage(param);
        }

        [ActionMethod]
        public static void ClickButton(string param)
        {
            Instantiator.MSDynamicsAXUI.ClickButton(param);
        }
        [ActionMethod]
        public static void ClickMenuItem(string param)
        {
            Instantiator.MSDynamicsAXUI.ClickMenuItem(param);
        }
        [ActionMethod]
        public static void ClickHyperlink(string param)
        {
            Instantiator.MSDynamicsAXUI.ClickHyperlink(param);
        }
        [ActionMethod]
        public static void SetText(string textName, string textvalue)
        {
            Instantiator.MSDynamicsAXUI.SetText(textName, textvalue);
        }
        [ActionMethod]
        public static void ClickCheckBox(string checkBoxName, bool value)
        {
            Instantiator.MSDynamicsAXUI.ClickCheckBox(checkBoxName, value);
        }

    }
}
