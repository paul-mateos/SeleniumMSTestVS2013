using System;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Security.Principal;

namespace RTA.Automation.CRM
{
    public class Impersonator : IDisposable
    {
        private WindowsImpersonationContext _wic;

        /// <summary>
        /// Begins impersonation with the given credentials, Logon type and Logon provider.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="domainName">Name of the domain.</param>
        /// <param name="password">The password. <see cref="System.String"/></param>
        /// <param name="logonType">Type of the logon.</param>
        /// <param name="logonProvider">The logon provider.</param>
        public Impersonator(string userName, string domainName, string password, WinApi.LogonType logonType, WinApi.LogonProvider logonProvider)
        {
            Impersonate(userName, domainName, password, logonType, logonProvider);
        }

        /// <summary>
        /// Begins impersonation with the given credentials.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="domainName">Name of the domain.</param>
        /// <param name="password">The password. <see cref="System.String"/></param>
        public Impersonator(string userName, string domainName, string password)
        {
            Impersonate(userName, domainName, password, WinApi.LogonType.Logon32LogonInteractive, WinApi.LogonProvider.Logon32ProviderDefault);
        }

        /// <summary>
        /// Initializes a new instance of the <see cref="Impersonator"/> class.
        /// </summary>
        public Impersonator()
        { }

        /// <summary>
        /// Performs application-defined tasks associated with freeing, releasing, or resetting unmanaged resources.
        /// </summary>
        public void Dispose()
        {
            UndoImpersonation();
        }

        /// <summary>
        /// Impersonates the specified user account.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="domainName">Name of the domain.</param>
        /// <param name="password">The password. <see cref="System.String"/></param>
        public void Impersonate(string userName, string domainName, string password)
        {
            Impersonate(userName, domainName, password, WinApi.LogonType.Logon32LogonInteractive, WinApi.LogonProvider.Logon32ProviderDefault);
        }

        /// <summary>
        /// Impersonates the specified user account.
        /// </summary>
        /// <param name="userName">Name of the user.</param>
        /// <param name="domainName">Name of the domain.</param>
        /// <param name="password">The password. <see cref="System.String"/></param>
        /// <param name="logonType">Type of the logon.</param>
        /// <param name="logonProvider">The logon provider.</param>
        public void Impersonate(string userName, string domainName, string password, WinApi.LogonType logonType, WinApi.LogonProvider logonProvider)
        {
            UndoImpersonation();

            var logonToken = IntPtr.Zero;
            var logonTokenDuplicate = IntPtr.Zero;

            try
            {
                // revert to the application pool identity, saving the identity of the current requestor
                _wic = WindowsIdentity.Impersonate(IntPtr.Zero);

                // do logon & impersonate
                if (WinApi.LogonUser(userName, domainName, password, (int)logonType, (int)logonProvider, ref logonToken) != 0)
                {
                    if (WinApi.DuplicateToken(logonToken, (int)WinApi.ImpersonationLevel.SecurityImpersonation, ref logonTokenDuplicate) != 0)
                    {
                        var wi = new WindowsIdentity(logonTokenDuplicate);
                        wi.Impersonate();
                    }
                    else
                    {
                        throw new Win32Exception(Marshal.GetLastWin32Error());
                    }
                }
                else
                {
                    throw new Win32Exception(Marshal.GetLastWin32Error());
                }
            }
            finally
            {
                if (logonToken != IntPtr.Zero)
                    WinApi.CloseHandle(logonToken);

                if (logonTokenDuplicate != IntPtr.Zero)
                    WinApi.CloseHandle(logonTokenDuplicate);
            }
        }

        /// <summary>
        /// Stops impersonation.
        /// </summary>
        private void UndoImpersonation()
        {
            // restore saved requestor identity
            if (_wic != null) _wic.Undo();
            _wic = null;
        }
    }
}
