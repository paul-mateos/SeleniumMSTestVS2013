using AutoItX3Lib;
using RTA.Automation.CRM.Environment;
using System;
using System.Windows.Automation;

namespace RTA.Automation.CRM.UI
{
    /**
     * This whole apporach needs further consideration - maybe using UI Automation, or ideally avoid the 
     * authentication window altogether - NTLM passthrough or some proxy solution maybe?
     */
    public class LoginDialog
    {

        private AutoItX3 AutoIt;
        public void Login(String username, String password)
        {
            AutoIt = new AutoItX3();
            switch (Properties.Settings.Default.BROWSER)
            {
                case BrowserType.Ie:
                    this.LoginIe(username, password);
                    break;
                case BrowserType.Chrome:
                    this.LoginChrome(username, password);
                    break;
                case BrowserType.Firefox:
                    this.LoginFirefox(username, password);
                    break;
            }
        }

        private void LoginFirefox(string username, string password)
        {
            throw new NotImplementedException();
        }

        private void LoginChrome(string username, string password)
        {
            throw new Exception("Cannot get Chrome Login to work successfully at present");
            /*
            string windowText = "Authentication Required";
            AutoIt.AutoItSetOption("WinTitleMatchMode", 2);
            AutoIt.WinWaitActive("",windowText, Properties.Settings.Default.SHORT_WAIT_SECONDS);
          
            if(AutoIt.WinExists("",windowText) == 1)
            {
                AutoIt.Send(username);
                AutoIt.Send("{TAB}");
                AutoIt.Send(password);
                AutoIt.Send("{ENTER}");

                AutoIt.WinWaitClose("", windowText);
            }
            */
        }

        private void LoginIe(string username, string password)
        {
            string windowName = "Windows Security";

            int found = AutoIt.WinWait(windowName, "", Properties.Settings.Default.SHORT_WAIT_SECONDS);
            if (found == 1)
            {
                AutoIt.WinActivate(windowName);
                AutoIt.Send(username);
                AutoIt.Send("{TAB}");
                AutoIt.Send(password);
                AutoIt.Send("{ENTER}");

                AutoIt.WinWaitClose(windowName);
            }
        }
    }
}
