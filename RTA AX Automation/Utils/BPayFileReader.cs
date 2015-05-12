using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Collections;

namespace RTA.Automation.AX.Utils
{
    class BPayFileReaderClass
    {
        public static string GetPaymentReference(string targetDirectory)
        {
            // Process the list of files found in the directory. 
            string[] fileEntries = Directory.GetFiles(@targetDirectory, "BPAY*.txt");
            if (fileEntries.Count() > 1)
            {
                throw new Exception(String.Format("There are to many files in the directory."));
            }
            else
            {

                StreamReader sr = new StreamReader(fileEntries.ElementAt(0));
                string[] lines;

                lines = sr.ReadToEnd().Split(Environment.NewLine.ToCharArray(), System.StringSplitOptions.RemoveEmptyEntries);
                if (!string.IsNullOrWhiteSpace(lines.ElementAt(0)))
                {
                    string[] ar = lines.ElementAt(3).Split(',');
                    return ar.ElementAt(4);

                }
                else
                {
                    throw new Exception(String.Format("File is empty"));
                }

            }
        }

        public static string GetPaymentReference1File(string fileLocation)
        {
            // Process the file found in the directory. 
            
                StreamReader sr = new StreamReader(fileLocation);
                string[] lines;

                lines = sr.ReadToEnd().Split(Environment.NewLine.ToCharArray(), System.StringSplitOptions.RemoveEmptyEntries);
                if (!string.IsNullOrWhiteSpace(lines.ElementAt(0)))
                {
                    string[] ar = lines.ElementAt(3).Split(',');
                    return ar.ElementAt(4);

                }
                else
                {
                    throw new Exception(String.Format("File is empty"));
                }
            
       

        }

        public static string GetTenancyRequestReference(string targetDirectory)
        {
            // Process the list of files found in the directory. 
            string[] fileEntries = Directory.GetFiles(@targetDirectory, "BPAY*.txt");
            if (fileEntries.Count() > 1)
            {
                throw new Exception(String.Format("There are to many files in the directory."));
            }
            else
            {

                StreamReader sr = new StreamReader(fileEntries.ElementAt(0));
                string[] lines;

                lines = sr.ReadToEnd().Split(Environment.NewLine.ToCharArray(), System.StringSplitOptions.RemoveEmptyEntries);
                if (!string.IsNullOrWhiteSpace(lines.ElementAt(0)))
                {
                    string[] ar = lines.ElementAt(0).Split(',');
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