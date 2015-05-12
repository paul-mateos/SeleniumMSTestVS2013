using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RTA.Automation.CRM.Environment
{
    public class TestEnvironment
    {
        public static TestEnvironment GetTestEnvironment() 
        {
            switch (Properties.Settings.Default.ENVIRONMENT)
            {
                case EnvironmentType.SystemTest:
                    return GetSystemTestEnvironment();
                case EnvironmentType.SIT:
                    return GetSitEnvironment();
                case EnvironmentType.SME:
                    return GetSMETestEnvironment();
                case EnvironmentType.ModelOffice:
                    return GetModelOfficeTestEnvironment();
                case EnvironmentType.IRSIT:
                    return GetIRSitTestEnvironment();
                default:
                    throw new ArgumentException("Invalid ENVIRONMENT Setting has been used");
            }
        }

        
        private static TestEnvironment GetSitEnvironment()
        {
            //this could be read in from an external source such as database, file etc.
            List<User> users = new List<User>();
            users.Add(new User(SecurityRole.Default, "imstestu12", "Password12"));
            users.Add(new User(SecurityRole.Investigations, "imstestu04", "Password4"));
            users.Add(new User(SecurityRole.SystemAdministrator, "imstestu12", "Password12"));
            users.Add(new User(SecurityRole.RBSOfficer, "imstestu13", "Password13"));
            users.Add(new User(SecurityRole.InvestigationsManager, "imstestu05", "Password5"));
            users.Add(new User(SecurityRole.InvestigationsOfficer, "imstestu02", "Password2")); // This user is Investigation Support Officer
            users.Add(new User(SecurityRole.GeneralStaff, "imstestu01", "Password1"));
            users.Add(new User(SecurityRole.InvestigationsBusinessAdmin, "imstestu06", "Password6"));
            users.Add(new User(SecurityRole.RBSClaimsOfficer, "imstestu21", "Password21"));
            users.Add(new User(SecurityRole.ESOForPES, "IMSTestU07", "Password7"));
            users.Add(new User(SecurityRole.ExecutiveManagerForPES, "IMSTestU08", "Password8"));
            users.Add(new User(SecurityRole.RecordKeepingOfficers, "IMSTestU09", "Password9"));
            users.Add(new User(SecurityRole.ResearchOfficers, "IMSTestU10", "Password10"));
            users.Add(new User(SecurityRole.IMSBusinessSupportStaff, "IMSTestU11", "Password11"));
            users.Add(new User(SecurityRole.InvestigationOfficer, "IMSTestU03", "Password3")); // This user is Investigation Officer

            //users.Add(new User(SecurityRole.SystemAdministrator, "admin", "admin"));
            return new TestEnvironment("http://rtacrmuat/MSCRMRTA05/main.aspx", users);
        }

        private static TestEnvironment GetSystemTestEnvironment()
        {
            //this could be read in from an external source such as database, file etc.
            List<User> users = new List<User>();
            users.Add(new User(SecurityRole.Default, "imstestu12", "Password12"));
            users.Add(new User(SecurityRole.Investigations, "imstestu04", "Password4"));
            users.Add(new User(SecurityRole.SystemAdministrator, "imstestu12", "Password12"));
            users.Add(new User(SecurityRole.RBSOfficer, "imstestu13", "Password13"));
            users.Add(new User(SecurityRole.InvestigationsManager, "imstestu05", "Password5"));
            users.Add(new User(SecurityRole.InvestigationsOfficer, "imstestu02", "Password2")); // This user is Investigation Support Officer
            users.Add(new User(SecurityRole.GeneralStaff, "imstestu01", "Password1"));
            users.Add(new User(SecurityRole.InvestigationsBusinessAdmin, "imstestu06", "Password6"));
            users.Add(new User(SecurityRole.RBSClaimsOfficer, "imstestu21", "Password21"));
            users.Add(new User(SecurityRole.ESOForPES, "IMSTestU07", "Password7"));
            users.Add(new User(SecurityRole.ExecutiveManagerForPES, "IMSTestU08", "Password8"));
            users.Add(new User(SecurityRole.RecordKeepingOfficers, "IMSTestU09", "Password9"));
            users.Add(new User(SecurityRole.ResearchOfficers, "IMSTestU10", "Password10"));
            users.Add(new User(SecurityRole.IMSBusinessSupportStaff, "IMSTestU11", "Password11"));
            users.Add(new User(SecurityRole.InvestigationOfficer, "IMSTestU03", "Password3")); // This user is Investigation Officer

            //users.Add(new User(SecurityRole.SystemAdministrator, "admin", "admin"));
            return new TestEnvironment("http://srcrm51-te/MSCRMRTA08/main.aspx", users);
        }


        private static TestEnvironment GetSMETestEnvironment()
        {
            //this could be read in from an external source such as database, file etc.
            List<User> users = new List<User>();
            users.Add(new User(SecurityRole.Default, "imstestu12", "Password12"));
            users.Add(new User(SecurityRole.Investigations, "imstestu04", "Password4"));
            users.Add(new User(SecurityRole.SystemAdministrator, "imstestu12", "Password12"));
            users.Add(new User(SecurityRole.RBSOfficer, "imstestu13", "Password13"));
            users.Add(new User(SecurityRole.InvestigationsManager, "imstestu05", "Password5"));
            users.Add(new User(SecurityRole.InvestigationsOfficer, "imstestu02", "Password2")); // This user is Investigation Support Officer
            users.Add(new User(SecurityRole.GeneralStaff, "imstestu01", "Password1"));
            users.Add(new User(SecurityRole.InvestigationsBusinessAdmin, "imstestu06", "Password6"));
            users.Add(new User(SecurityRole.RBSClaimsOfficer, "imstestu21", "Password21"));
            users.Add(new User(SecurityRole.ESOForPES, "IMSTestU07", "Password7"));
            users.Add(new User(SecurityRole.ExecutiveManagerForPES, "IMSTestU08", "Password8"));
            users.Add(new User(SecurityRole.RecordKeepingOfficers, "IMSTestU09", "Password9"));
            users.Add(new User(SecurityRole.ResearchOfficers, "IMSTestU10", "Password10"));
            users.Add(new User(SecurityRole.IMSBusinessSupportStaff, "IMSTestU11", "Password11"));
            users.Add(new User(SecurityRole.InvestigationOfficer, "IMSTestU03", "Password3")); // This user is Investigation Officer

            //users.Add(new User(SecurityRole.SystemAdministrator, "admin", "admin"));
            return new TestEnvironment("http://srcrmsme63-ua:5555/SME/main.aspx", users);
        }

        private static TestEnvironment GetModelOfficeTestEnvironment()
        {
            //this could be read in from an external source such as database, file etc.
            List<User> users = new List<User>();
            users.Add(new User(SecurityRole.Default, "imstestu12", "Password12"));
            users.Add(new User(SecurityRole.Investigations, "imstestu04", "Password4"));
            users.Add(new User(SecurityRole.SystemAdministrator, "imstestu12", "Password12"));
            users.Add(new User(SecurityRole.RBSOfficer, "imstestu13", "Password13"));
            users.Add(new User(SecurityRole.InvestigationsManager, "imstestu05", "Password5"));
            users.Add(new User(SecurityRole.InvestigationsOfficer, "imstestu02", "Password2")); // This user is Investigation Support Officer
            users.Add(new User(SecurityRole.GeneralStaff, "imstestu01", "Password1"));
            users.Add(new User(SecurityRole.InvestigationsBusinessAdmin, "imstestu06", "Password6"));
            users.Add(new User(SecurityRole.RBSClaimsOfficer, "imstestu21", "Password21"));
            users.Add(new User(SecurityRole.ESOForPES, "IMSTestU07", "Password7"));
            users.Add(new User(SecurityRole.ExecutiveManagerForPES, "IMSTestU08", "Password8"));
            users.Add(new User(SecurityRole.RecordKeepingOfficers, "IMSTestU09", "Password9"));
            users.Add(new User(SecurityRole.ResearchOfficers, "IMSTestU10", "Password10"));
            users.Add(new User(SecurityRole.IMSBusinessSupportStaff, "IMSTestU11", "Password11"));
            users.Add(new User(SecurityRole.InvestigationOfficer, "IMSTestU03", "Password3")); // This user is Investigation Officer

            //users.Add(new User(SecurityRole.SystemAdministrator, "admin", "admin"));
            return new TestEnvironment("http://srcrm01-pr/MSCRMRTA/main.aspx", users);
        }

        private static TestEnvironment GetIRSitTestEnvironment()
        {
            //this could be read in from an external source such as database, file etc.
            List<User> users = new List<User>();
            users.Add(new User(SecurityRole.Default, "imstestu12", "Password12"));
            users.Add(new User(SecurityRole.Investigations, "imstestu04", "Password4"));
            users.Add(new User(SecurityRole.SystemAdministrator, "imstestu12", "Password12"));
            users.Add(new User(SecurityRole.RBSOfficer, "imstestu13", "Password13"));
            users.Add(new User(SecurityRole.InvestigationsManager, "imstestu05", "Password5"));
            users.Add(new User(SecurityRole.InvestigationsOfficer, "imstestu02", "Password2")); // This user is Investigation Support Officer
            users.Add(new User(SecurityRole.GeneralStaff, "imstestu01", "Password1"));
            users.Add(new User(SecurityRole.InvestigationsBusinessAdmin, "imstestu06", "Password6"));
            users.Add(new User(SecurityRole.RBSClaimsOfficer, "imstestu21", "Password21"));
            users.Add(new User(SecurityRole.ESOForPES, "IMSTestU07", "Password7"));
            users.Add(new User(SecurityRole.ExecutiveManagerForPES, "IMSTestU08", "Password8"));
            users.Add(new User(SecurityRole.RecordKeepingOfficers, "IMSTestU09", "Password9"));
            users.Add(new User(SecurityRole.ResearchOfficers, "IMSTestU10", "Password10"));
            users.Add(new User(SecurityRole.IMSBusinessSupportStaff, "IMSTestU11", "Password11"));
            users.Add(new User(SecurityRole.InvestigationOfficer, "IMSTestU03", "Password3")); // This user is Investigation Officer

            //users.Add(new User(SecurityRole.SystemAdministrator, "admin", "admin"));
            return new TestEnvironment("http://rtacrmuat/MSCRMRTA/main.aspx", users);
        }
        public string Url
        {
            get;
            private set;
        }

       
        public TestEnvironment(string url, List<User> users) {
            this.Url = url;
            this.Users = users;
        }

        public List<User> Users { get; set; }

       
        public User GetUser(SecurityRole roleUse)
        {
            

            for (int i=0;i<this.Users.Count;i++)
            {
                if (this.Users[i].Role == roleUse)
                {
                    return this.Users[i];
                }
            }

            throw new NotImplementedException();
        }
    }
}
