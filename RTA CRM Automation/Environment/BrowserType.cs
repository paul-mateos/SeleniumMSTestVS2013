using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTA.Automation.CRM.Environment
{
    public enum BrowserType
    {
        [Description("firefox")]
        Firefox = 0,
        [Description("chrome")]
        Chrome = 1,
        [Description("ie")]
        Ie = 2
    }
}
