using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTA.Automation.CRM.Utils
{
    class BAI2FileCreator
    {
        public static string bAI2FileCreator(string referenceNumber, string amount)
        {
            string dateValue = DateTime.Today.ToString("yyMMdd");
            string timeValue = DateTime.Now.ToString("HHmm");

            string Line1 = "01,CBA,RTAUTHQL," + dateValue + "," + timeValue + ",1,,,2/";
            string Line2 = "02,,CBA,1," + dateValue + ",,AUD,2/";
            string Line3 = "03,401310006413,,015,,,,100," + Convert.ToInt32(amount) * 100 + ",1,,400,0,0,,900,000,,,901,000,,,902,000,,,903,000,,,904,,,,905,,,/";
            string Line4 = "16,399," + Convert.ToInt32(amount) * 100 + ",,MIS,," + referenceNumber + "/";
            string Line5 = "49," + ((Convert.ToInt32(amount) * 2) * 100) + ",3/";
            string Line6 = "98," + ((Convert.ToInt32(amount) * 2) * 100) + ",5/";
            string Line7 = "99," + ((Convert.ToInt32(amount) * 2) * 100) + ",1,7/";
            // Create a string array that consists of three lines. 
            string[] lines = { Line1, Line2, Line3, Line4, Line5, Line6, Line7};
            Random random = new Random();
            int randomNum = random.Next(1000, 9999);
            string filelocation = @"P:\Dynamics AX\Bank files\Bank Statements\Paul\BAI2-AUTOMATION-" + dateValue + "-" + referenceNumber + "-" + randomNum + ".txt";
            System.IO.File.WriteAllLines(filelocation, lines);
            return filelocation;
        }

    }
}
