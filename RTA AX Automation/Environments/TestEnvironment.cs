using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RTA.Automation.AX.Environments
{
    public class TestEnvironment
    {
        public static string GetTestEnvironment() 
        {
            switch (Properties.Settings.Default.ENVIRONMENT)
            {
                case EnvironmentType.SystemTest:
                    return @"P:\Dynamics AX\test 32.axc";
                    //return @"Microsoft Dynamics AX\60\Client\Bin\test 32.axc";
                case EnvironmentType.SIT:
                    return @"P:\Dynamics AX\sit 32.axc";
                    //return @"Microsoft Dynamics AX\60\Client\Bin\sit 32.axc";
                case EnvironmentType.SME:
                    return @"P:\Dynamics AX\sme 32.axc";
                    //return @"Microsoft Dynamics AX\60\Client\Bin\sme 32.axc";
                case EnvironmentType.IRSIT:
                    //return @"P:\Dynamics AX\sme 32.axc";
                    throw new Exception("No IRSIT Available");
                //return @"Microsoft Dynamics AX\60\Client\Bin\sme 32.axc";
                default:
                    throw new ArgumentException("Invalid ENVIRONMENT Setting has been used");
            }
        }

       
    }
}
