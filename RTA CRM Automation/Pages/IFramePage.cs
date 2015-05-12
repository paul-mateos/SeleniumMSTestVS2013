using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace RTA.Automation.CRM.Pages
{
    public abstract class IFramePage : BasePage
    {
        protected string frame;
       

        public IFramePage(IWebDriver driver, string frame)
            : base(driver)
        {
            int retries = 3;
            int Count = 0;
            while (Count < retries)
            {
                try
                {
                    this.frame = frame;
                    driver.SwitchTo().Frame(frame);
                    break;
                }
                catch
                {
                    driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
                    Count++;
                }
            }
        }

        public IFramePage(IWebDriver driver)
            : base(driver)
        {
            driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(2));
            this.frame = null;
        }
    }
}
