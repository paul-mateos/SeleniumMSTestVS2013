using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTA.Automation.AX.Utils
{
    class BAI2FileCreator
    {
        public static string bAI2InvalidRefFileCreator()
        {
            string referenceNumber1 = "ILY 3001 1205 CLOSED";
            Random random = new Random(); 
            int randomNum = random.Next(30000000, 39999999);
            string referenceNumber2 = randomNum.ToString();
            string amount = "500";
            string amountTotal = "1000";
            string dateValue = DateTime.Today.ToString("yyMMdd");
            string timeValue = DateTime.Now.ToString("HHmm");

            string Line1 = "01,CBA,RTAUTHQL," + dateValue + "," + timeValue + ",1,,,2/";
            string Line2 = "02,,CBA,1," + dateValue + ",,AUD,2/";
            string Line3 = "03,401310006413,,015,,,,100," + Convert.ToInt32(amountTotal) * 100 + ",2,,400,0,0,,900,000,,,901,000,,,902,000,,,903,000,,,904,,,,905,,,/";
            string Line4 = "16,399," + Convert.ToInt32(amount) * 100 + ",,MIS,," + referenceNumber1 + "/";
            string Line5 = "16,399," + Convert.ToInt32(amount) * 100 + ",,MIS,," + referenceNumber2 + "/";
            string Line6 = "49," + ((Convert.ToInt32(amountTotal) * 2) * 100) + ",4/";
            string Line7 = "98," + ((Convert.ToInt32(amountTotal) * 2) * 100) + ",6/";
            string Line8 = "99," + ((Convert.ToInt32(amountTotal) * 2) * 100) + ",1,8/";
            // Create a string array that consists of three lines. 
            string[] lines = { Line1, Line2, Line3, Line4, Line5, Line6, Line7, Line8};
            
            randomNum = random.Next(1000, 9999);
            string filelocation = @"P:\Dynamics AX\Bank files\Bank Statements\Paul\BAI2-AUTOMATION-" + dateValue + "-IGNORED and UNKNOWN-" + referenceNumber2 + ".txt";
            System.IO.File.WriteAllLines(filelocation, lines);
            return filelocation;
        }

        public static string bAI2UnknownCreator(int random, string amount, string dateValue, string timeValue)
        {
            string Line1 = "01,CBA,RTAUTHQL," + dateValue + "," + timeValue + ",1,,,2/";
            string Line2 = "02,,CBA,1," + dateValue + ",,AUD,2/";
            string Line3 = "03,401310006413,,015,,,,100," + Convert.ToInt32(amount) * 100 + ",1,,400,0,0,,900,000,,,901,000,,,902,000,,,903,000,,,904,,,,905,,,/";
            string Line4 = "16,399," + Convert.ToInt32(amount) * 100 + ",,MIS,," + random + "/";
            string Line5 = "49," + ((Convert.ToInt32(amount) * 2) * 100) + ",3/";
            string Line6 = "98," + ((Convert.ToInt32(amount) * 2) * 100) + ",5/";
            string Line7 = "99," + ((Convert.ToInt32(amount) * 2) * 100) + ",1,7/";
            // Create a string array that consists of three lines. 
            string[] lines = { Line1, Line2, Line3, Line4, Line5, Line6, Line7 };
            string filelocation = @"P:\Dynamics AX\Bank files\Bank Statements\Paul\BAI2-AUTOMATION-" + dateValue + "-" + random + ".txt";
            System.IO.File.WriteAllLines(filelocation, lines);
            return filelocation;
        }

       
    }
}
