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
    public class BPayFileExceptions
    {
       
        [ActionMethod]
        public static void ClickTab(string param)
        {
            //Instantiator.BPayFileExceptionsUI.ClickTabPage(param);
        }

        [ActionMethod]
        public static void ClickButton(string param)
        {
            //Instantiator.BPayFileExceptionsUI.ClickButton(param);
        }
        [ActionMethod]
        public static void ClickMenuItem(string param)
        {
            //Instantiator.BPayFileExceptionsUI.ClickMenuItem(param);
        }
        [ActionMethod]
        public static void ClickHyperlink(string param)
        {
            //Instantiator.BPayFileExceptionsUI.ClickHyperlink(param);
        }
        //[ActionMethod]
        //public static void SetText(string textName, string textvalue)
        //{
        //    Instantiator.BPayFileExceptionsUI.SetText(textName, textvalue);
        //}
        //[ActionMethod]
        //public static void ClickCheckBox(string checkBoxName, bool value)
        //{
        //    Instantiator.BPayFileExceptionsUI.ClickCheckBox(checkBoxName, value);
        //}
        [ActionMethod]
        public static bool FindTableCellValue(string searchValue)
        {
            //WinCell cellValue = Instantiator.BPayFileExceptionsUI.mUIWinTable.FindFirstCellWithValue("RTB-000331");
            bool returnValue = Instantiator.BPayFileExceptionsUI.FindTableCellValue(searchValue);
           // Instantiator.BPayFileExceptionsUI.table();
            return returnValue;
        }
               
        
    }
}
