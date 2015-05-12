using System;

namespace RTA.Automation.CRM
{
    [Serializable]
    public class SecurityCredentials
    {
        public string User { get; set; }
        public string Domain { get; set; }
        public byte[] EncryptedPassword { get; set; }
    }
}
