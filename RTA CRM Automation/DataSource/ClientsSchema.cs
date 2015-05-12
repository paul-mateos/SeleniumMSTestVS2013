using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTA.Automation.CRM.DataSource
{
    class ClientsSchema
    {
        
        public static int GetColumnIndex(string columnName)
        {
            switch (columnName)
            {
                case "TESTID":
                    return 1;
                case "DESCRIPTION":
                    return 2;
                case "CLIENT_NAME":
                    return 3;
                case "TITLE":
                    return 4;
                case "GIVEN_NAME":
                    return 5;
                case "MIDDLE_NAME":
                    return 6;
                case "FAMILY_NAME":
                    return 7;
                case "SUFFIX":
                    return 8;
                case "EMAIL":
                    return 9;
                case "ADDRESS":
                    return 10;
                case "ID_PROVIDED":
                    return 11;
                default:
                    throw new ArgumentException("Invalid Column Heading in Data Source Schema ");
            }
                
        }
    }
}
