using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTA.Automation.CRM.DataSource
{
    class TenancyRequestSchema
    {
        //[Flags]
        //public enum ColumnName
        //{
        //    TESTID = 1,
        //    DESCRIPTION = 2,
        //    IN_ID = 3
        //}

        public static int GetColumnIndex(string columnName)
        {
            switch (columnName)
            {
                case "TESTID":
                    return 1;
                case "DESCRIPTION":
                    return 2;
                case "REQUEST_TYPE":
                    return 3;
                case "RENTAL_PREMISES":
                    return 4;
                case "MANAGING_PARTY":
                    return 5;
                case "TENANCY":
                    return 6;
                case "TENANCY_TYPE":
                    return 7;
                case "MANAGEMENT_TYPE":
                    return 8;
                case "DWELLING_TYPE":
                    return 9;
                case "INITIAL_REQUEST_PARTY":
                    return 10;
                case "INITIAL_CONTRIBUTION":
                    return 11;
                case "WEEKLY_RENT":
                    return 12;
                case "AMOUNT_PAID_LODGEMENT":
                    return 13;
                case "STATUS_REASON":
                    return 14;
                case "LODGEMENT_TYPE":
                    return 15;
                case "TENANCY_START":
                    return 16;
                case "ANTICIPATED_END":
                    return 17;
                case "NO_ROOMS":
                    return 18;
                case "FUNDED_STATUS":
                    return 19;
                case "TR_NUMBER":
                    return 20;
                case "PAY_REF_NUMBER":
                    return 21;
                case "PAYMENT_TYPE":
                    return 22;
                case "BOND_REF":
                    return 23;
                case "OUTFILE":
                    return 24;
                case "MISC1":
                    return 25;
                case "MISC2":
                    return 26;
                default:
                    throw new ArgumentException("Invalid Column Heading in Data Source Schema ");
            }
                
        }
    }
}
