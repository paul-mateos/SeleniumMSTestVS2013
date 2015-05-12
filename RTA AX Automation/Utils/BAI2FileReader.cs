using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;

namespace RTA.Automation.AX.Utils
{
    class BAI2FileReaderClass
    {
        public static string GetPaymentReference(string targetDirectory)
        {
            // Process the list of files found in the directory. 
            //string[] fileEntries = Directory.GetFiles(@targetDirectory, "BAI2*.txt");
            //if (fileEntries.Count() > 1)
            //{
            //    throw new Exception(String.Format("There are to many files in the directory."));
            //}
            //else
            //{

                StreamReader sr = new StreamReader(targetDirectory);//fileEntries.ElementAt(0));
                string[] lines;

                lines = sr.ReadToEnd().Split(Environment.NewLine.ToCharArray(), System.StringSplitOptions.RemoveEmptyEntries);
                if (!string.IsNullOrWhiteSpace(lines.ElementAt(0)))
                {
                    string[] ar = lines.ElementAt(3).Split(',');
                    return ar.ElementAt(6);

                }
                else
                {
                    throw new Exception(String.Format("File is empty"));
                }

            //}
        }

        
        
    }
    
   

}