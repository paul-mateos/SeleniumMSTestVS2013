using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTA.Automation.AX.Environments
{
    public enum EnvironmentType
    {
        [Description("SystemTest")]
        SystemTest = 0,
        [Description("SIT")]
        SIT = 1,
        [Description("SME")]
        SME = 2,
        [Description("IRSIT")]
        IRSIT = 3,
    }
}
