using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;

namespace RTA.Automation.AX.Utils
{
    class EFTFileReaderClass
    {

        public static string GetPaymentReference(string targetDirectory)
        {
            // Process the list of files found in the directory. 
            string[] fileEntries = Directory.GetFiles(@targetDirectory, "EFT*.txt");
            if (fileEntries.Count() > 1)
            {
                throw new Exception(String.Format("There are to many files in the directory."));
            }
            else
            {

                StreamReader sr = new StreamReader(fileEntries.ElementAt(0));
                string line;
                line = sr.ReadLine();
                if (!string.IsNullOrWhiteSpace(line))
                {
                    string[] ar = line.Split(',');            
                    return ar.ElementAt(5);
                                    
                }else
                {
                    throw new Exception(String.Format("File is empty"));
                }
                     
            }
        }

        public static string GetTenancyRequestReference(string targetDirectory)
        {
            // Process the list of files found in the directory. 
            string[] fileEntries = Directory.GetFiles(@targetDirectory, "EFT*.txt");
            if (fileEntries.Count() > 1)
            {
                throw new Exception(String.Format("There are to many files in the directory."));
            }
            else
            {

                StreamReader sr = new StreamReader(fileEntries.ElementAt(0));
                string line;
                line = sr.ReadLine();
                if (!string.IsNullOrWhiteSpace(line))
                {
                    string[] ar = line.Split(',');
                    return ar.ElementAt(2);

                }
                else
                {
                    throw new Exception(String.Format("File is empty"));
                }

            }
        }
    
     }
   

}