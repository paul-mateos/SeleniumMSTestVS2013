using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTA.Automation.CRM.Utils
{
    class EFTFileCreator
    {
        public static string eFTFileCreator(string tenancyrequest, string referenceNumber)
        {
            string dateValue = DateTime.Today.ToString("dd/MM/yyyy");
            string dateValue2 = DateTime.Today.ToString("yyyyMMdd");

            string Line1 = "401310006413," + dateValue + "," + tenancyrequest + ",1000.00,CR," + referenceNumber + ",732-299   606060,TEST EFT TENANCY,";
            string Line2 = "GRANDTOTAL," + dateValue + ",TRANS,1,CR AMT,1000.00,DR AMT,0.00,";
            // Create a string array that consists of three lines. 
            string[] lines = { Line1, Line2};
            Random random = new Random();
            int randomNum = random.Next(1000, 9999);
            System.IO.File.WriteAllLines(@"P:\Dynamics AX\Bank files\EFT\Paul\EFT-AUTOMATION-" + dateValue2 + "-" + referenceNumber + "-" + randomNum + ".txt", lines);
            return @"P:\Dynamics AX\Bank files\EFT\Paul\EFT-AUTOMATION-" + dateValue2 + "-" + referenceNumber + "-" + randomNum + ".txt";
        }

    }
}
