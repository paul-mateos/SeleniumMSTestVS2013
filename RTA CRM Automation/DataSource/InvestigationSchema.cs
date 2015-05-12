using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTA.Automation.CRM.DataSource
{
    [Flags]
    public enum ColumnName
    {
        TESTID = 1,
        DESCRIPTION = 2,
        IN_ID = 3,
        CLIENT_NAME = 4
    }

    class InvestigationSchema
    {
        public static int GetColumnIndex(ColumnName columnName)
        {
            switch (columnName)
            {
                case ColumnName.TESTID:
                    return 1;
                case ColumnName.DESCRIPTION:
                    return 2;
                case ColumnName.IN_ID:
                    return 3;
                case ColumnName.CLIENT_NAME:
                    return 4;
                default:
                    throw new ArgumentException("Invalid Column Heading in Data Source Schema ");
            }
                
        }
    }
}
