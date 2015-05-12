using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RTA.Automation.CRM.Environment
{
    public class User
    {

        public User(SecurityRole role, string id, string password)
        {
            this.Role = role;
            this.Id = id;
            this.Password = password;
        }
        
        public SecurityRole Role
        {
            get;
            private set;
        }

        public string Id
        {
            get;
            private set;
        }

        public string Password
        {
            get;
            private set;
        }


    }
}
