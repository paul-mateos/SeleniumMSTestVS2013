using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RTA.Automation.CRM.Pages
{


    public abstract class BasePage
    {
        protected IWebDriver driver = null;

        

        public BasePage(IWebDriver driver)
        {
            
            this.driver = driver;
         
            driver.SwitchTo().DefaultContent();
        }


    }
}
