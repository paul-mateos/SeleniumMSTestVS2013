using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTA.Automation.AX.Utils
{
    class BPayFileCreator
    {
       
        public static string bPayUnknownClientFileCreator(string dateValue, int randomNum)
        {

            string dateTimeValue = dateValue + DateTime.Now.ToString("hhmmss");
            string Line1 = "01,CBABPAY,UnknownClient," + dateValue + ",0110,1,,,2/";
            string Line2 = "02,5793,CBA,1," + dateValue + ",,,3/";
            string Line3 = "03,401310041964,,231,70000,101,,250,,0,,550,0,0,/";
            string Line4 = "30,399,70000,0," + dateTimeValue + ",CBA201409110759258765,0,05,001," + dateValue + ",144615,004,,,,,,,,,/";
            string Line5 = "49,140000,3/";
            string Line6 = "98,25260800,1,105/";
            string Line7 = "99,25260800,1,107/";


            // Create a string array that consists of three lines. 
            string[] lines = { Line1, Line2, Line3, Line4, Line5, Line6, Line7 };
            //Random random = new Random();
            //int randomNum = random.Next(1000, 9999);
            string fileLocation = @"P:\Dynamics AX\Bank files\Bpay\Paul\BPAY-AUTOMATION-" + dateValue + "-" + randomNum + ".txt";
            System.IO.File.WriteAllLines(fileLocation, lines);
            return fileLocation;

        }

       
    }
}
