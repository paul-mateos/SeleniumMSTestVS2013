using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTA.Automation.CRM.Utils
{
    class SwitchBrowsers
    {
        public void SwitchBowsers(IWebDriver driver)
        {
            //*****************This needs to be moved out of here********************************************
            string NewWindow = ""; //prepares for the new window handle
            string BaseWindow = "";
            ReadOnlyCollection<string> handles = null;
            handles = driver.WindowHandles;
            foreach (string handle in handles)
            {
                var Handles = handle;
                if (BaseWindow != handle)
                {
                    NewWindow = handle;

                    driver = driver.SwitchTo().Window(NewWindow);
                    break;
                }
            }

        }
    }
}
