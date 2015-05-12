using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTA.Automation.CRM.Environment
{
    public enum SecurityRole
    {
        [Description("Default")]
        Default = 0,
        [Description("SystemAdministrator")]
        SystemAdministrator = 1,
        [Description("Investigations")]
        Investigations = 2,
        [Description("RBSOfficer")]
        RBSOfficer = 3,
        [Description("InvestigationsManager")]
        InvestigationsManager = 4,
        [Description("InvestigationsOfficer")]
        InvestigationsOfficer = 5,
        [Description("GeneralStaff")]
        GeneralStaff = 6,
        [Description("RBSClaimsOfficer")]
        RBSClaimsOfficer = 7,
        [Description("InvestigationsBusinessAdmin")]
        InvestigationsBusinessAdmin = 8,   
        [Description("ESOForPES")]
        ESOForPES = 9, 
        [Description("ExecutiveManagerForPES")]
        ExecutiveManagerForPES = 10,
        [Description("RecordKeepingOfficers")]
        RecordKeepingOfficers = 11,
        [Description("ResearchOfficers")]
        ResearchOfficers = 12,
        [Description("IMSBusinessSupportStaff")]
        IMSBusinessSupportStaff = 13,
        [Description("InvestigationOfficer")]
        InvestigationOfficer = 14
    }
}
