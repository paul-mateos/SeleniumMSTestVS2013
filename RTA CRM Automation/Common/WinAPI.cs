using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using Microsoft.Win32.SafeHandles;

namespace RTA.Automation.CRM
{
    /// <summary>
    /// Helper class for Win32 API
    /// </summary>
    public static class WinApi
    {
        private static readonly byte[] _s_k;
        private static readonly byte[] _s_v;

        #region Vars

        internal const int HwndBroadcast = 0xffff;
        internal const int SwShownormal = 1;
        internal const int OsAnyServer = 29;

        public const int ReadControl = 0x00020000;
        public const int StandardRightsRequired = 0x000F0000;
        public const int StandardRightsRead = ReadControl;
        public const int StandardRightsWrite = ReadControl;
        public const int StandardRightsExecute = ReadControl;
        public const int StandardRightsAll = 0x001F0000;
        public const int SpecificRightsAll = 0x0000FFFF;
        public const int TokenAssignPrimary = 0x0001;
        public const int TokenDuplicate = 0x0002;
        public const int TokenImpersonate = 0x0004;
        public const int TokenQuery = 0x0008;
        public const int TokenQuerySource = 0x0010;
        public const int TokenAdjustPrivileges = 0x0020;
        public const int TokenAdjustGroups = 0x0040;
        public const int TokenAdjustDefault = 0x0080;
        public const int TokenAdjustSessionid = 0x0100;
        public const int TokenAllAccessP = (StandardRightsRequired | TokenAssignPrimary | TokenDuplicate | TokenImpersonate | TokenQuery | TokenQuerySource | TokenAdjustPrivileges | TokenAdjustGroups | TokenAdjustDefault);
        public const int TokenAllAccess = TokenAllAccessP | TokenAdjustSessionid;
        public const int TokenRead = StandardRightsRead | TokenQuery;
        public const int TokenWrite = StandardRightsWrite | TokenAdjustPrivileges | TokenAdjustGroups | TokenAdjustDefault;
        public const int TokenExecute = StandardRightsExecute;
        public const uint MaximumAllowed = 0x2000000;
        public const int CreateNewProcessGroup = 0x00000200;
        public const int CreateUnicodeEnvironment = 0x00000400;
        public const int IdlePriorityClass = 0x40;
        public const int NormalPriorityClass = 0x20;
        public const int HighPriorityClass = 0x80;
        public const int RealtimePriorityClass = 0x100;
        public const int CreateNewConsole = 0x00000010;
        public const string SeDebugName = "SeDebugPrivilege";
        public const string SeRestoreName = "SeRestorePrivilege";
        public const string SeBackupName = "SeBackupPrivilege";
        public const int SePrivilegeEnabled = 0x0002;
        public const int ErrorNotAllAssigned = 1300;
        public const UInt32 Infinite = 0xFFFFFFFF;
        public static int InvalidHandleValue = -1;

        [StructLayout(LayoutKind.Sequential)]
        public struct SecurityAttributes
        {
            public int Length;
            public IntPtr lpSecurityDescriptor;
            public bool bInheritHandle;
        }

        public enum TokenInformationClass
        {
            TokenUser = 1,
            TokenGroups,
            TokenPrivileges,
            TokenOwner,
            TokenPrimaryGroup,
            TokenDefaultDacl,
            TokenSource,
            TokenType,
            TokenImpersonationLevel,
            TokenStatistics,
            TokenRestrictedSids,
            TokenSessionId,
            TokenGroupsAndPrivileges,
            TokenSessionReference,
            TokenSandBoxInert,
            TokenAuditPolicy,
            TokenOrigin,
            MaxTokenInfoClass
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct Startupinfo
        {
            public int cb;
            public String lpReserved;
            public String lpDesktop;
            public String lpTitle;
            public uint dwX;
            public uint dwY;
            public uint dwXSize;
            public uint dwYSize;
            public uint dwXCountChars;
            public uint dwYCountChars;
            public uint dwFillAttribute;
            public uint dwFlags;
            public short wShowWindow;
            public short cbReserved2;
            public IntPtr lpReserved2;
            public IntPtr hStdInput;
            public IntPtr hStdOutput;
            public IntPtr hStdError;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct ProcessInformation
        {
            public IntPtr hProcess;
            public IntPtr hThread;
            public uint dwProcessId;
            public uint dwThreadId;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct Luid
        {
            public int LowPart;
            public int HighPart;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal struct LuidAndAtributes
        {
            public Luid Luid;
            public int Attributes;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct TokenPrivileges
        {
            internal int PrivilegeCount;
            [MarshalAs(UnmanagedType.ByValArray, SizeConst = 3)]
            internal int[] Privileges;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct Processentry32
        {
            public uint dwSize;
            public uint cntUsage;
            public uint th32ProcessID;
            public IntPtr th32DefaultHeapID;
            public uint th32ModuleID;
            public uint cntThreads;
            public uint th32ParentProcessID;
            public int pcPriClassBase;
            public uint dwFlags;
            [MarshalAs(UnmanagedType.ByValTStr, SizeConst = 260)]
            public string szExeFile;
        }

        [StructLayout(LayoutKind.Sequential)]
        public struct WtsSessionInfo
        {
            public Int32 SessionID;

            [MarshalAs(UnmanagedType.LPStr)]
            public String pWinStationName;

            public WtsConnectState State;
        }

        public enum WtsInfoClass
        {
            WtsInitialProgram,
            WtsApplicationName,
            WtsWorkingDirectory,
            WtsoemId,
            WtsSessionId,
            WtsUserName,
            WtsWinStationName,
            WtsDomainName,
            WtsConnectState,
            WtsClientBuildNumber,
            WtsClientName,
            WtsClientDirectory,
            WtsClientProductId,
            WtsClientHardwareId,
            WtsClientAddress,
            WtsClientDisplay,
            WtsClientProtocolType
        }

        public enum WtsConnectState
        {
            WtsActive,
            WtsConnected,
            WtsConnectQuery,
            WtsShadow,
            WtsDisconnected,
            WtsIdle,
            WtsListen,
            WtsReset,
            WtsDown,
            WtsInit
        }

        public enum LogonType
        {
            Logon32LogonInteractive = 2,
            Logon32LogonNetwork = 3,
            Logon32LogonBatch = 4,
            Logon32LogonService = 5,
            Logon32LogonUnlock = 7,
            Logon32LogonNetworkCleartext = 8, // Win2K or higher
            Logon32LogonNewCredentials = 9 // Win2K or higher
        };

        public enum LogonProvider
        {
            Logon32ProviderDefault = 0,
            Logon32ProviderWinnt35 = 1,
            Logon32ProviderWinnt40 = 2,
            Logon32ProviderWinnt50 = 3
        };

        public enum ImpersonationLevel
        {
            SecurityAnonymous = 0,
            SecurityIdentification = 1,
            SecurityImpersonation = 2,
            SecurityDelegation = 3
        }

        public struct RunAsResult
        {
            public int ProcessId { get; set; }
            public int ExitCode { get; set; }
        }
        
        #endregion

        #region Externs

        [DllImport("user32")]
        public static extern int RegisterWindowMessage(string message);

        [DllImport("user32")]
        public static extern bool PostMessage(IntPtr hwnd, int msg, IntPtr wparam, IntPtr lparam);

        [DllImport("user32.dll")]
        public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("kernel32.dll")]
        public static extern int WTSGetActiveConsoleSessionId();

        [DllImport("Wtsapi32.dll")]
        public static extern bool WTSQueryUserToken(uint sessionId, ref IntPtr phToken);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool LookupPrivilegeValue(IntPtr lpSystemName, string lpname, [MarshalAs(UnmanagedType.Struct)] ref Luid lpLuid);

        [DllImport("advapi32.dll", EntryPoint = "CreateProcessAsUser", SetLastError = true, CharSet = CharSet.Ansi, CallingConvention = CallingConvention.StdCall)]
        public extern static bool CreateProcessAsUser(IntPtr hToken, string lpApplicationName, string lpCommandLine, IntPtr lpProcessAttributes, IntPtr lpThreadAttributes, bool bInheritHandle, int dwCreationFlags, IntPtr lpEnvironment, string lpCurrentDirectory, ref Startupinfo lpStartupInfo, out ProcessInformation lpProcessInformation);

        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public extern static int DuplicateToken(IntPtr existingTokenHandle, int securityImpersonationLevel, ref IntPtr duplicateTokenHandle);

        [DllImport("advapi32.dll", EntryPoint = "DuplicateTokenEx")]
        public extern static bool DuplicateTokenEx(IntPtr existingTokenHandle, uint dwDesiredAccess, ref SecurityAttributes lpThreadAttributes, int tokenType, int impersonationLevel, ref IntPtr duplicateTokenHandle);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool AdjustTokenPrivileges(IntPtr tokenHandle, bool disableAllPrivileges, ref TokenPrivileges newState, int bufferLength, IntPtr previousState, IntPtr returnLength);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern bool SetTokenInformation(IntPtr tokenHandle, TokenInformationClass tokenInformationClass, uint tokenInformation, uint tokenInformationLength);

        [DllImport("userenv.dll", SetLastError = true)]
        public static extern bool CreateEnvironmentBlock(out IntPtr lpEnvironment, IntPtr hToken, bool bInherit);

        [DllImport("userenv.dll", SetLastError = true)]
        public static extern bool DestroyEnvironmentBlock(IntPtr lpEnvironment);

        [DllImport("wtsapi32.dll")]
        static extern IntPtr WTSOpenServer([MarshalAs(UnmanagedType.LPStr)] String pServerName);

        [DllImport("wtsapi32.dll")]
        static extern void WTSCloseServer(IntPtr hServer);

        [DllImport("wtsapi32.dll")]
        static extern Int32 WTSEnumerateSessions(
            IntPtr hServer,
            [MarshalAs(UnmanagedType.U4)] Int32 reserved,
            [MarshalAs(UnmanagedType.U4)] Int32 version,
            ref IntPtr ppSessionInfo,
            [MarshalAs(UnmanagedType.U4)] ref Int32 pCount);

        [DllImport("wtsapi32.dll")]
        static extern void WTSFreeMemory(IntPtr pMemory);

        [DllImport("Wtsapi32.dll")]
        static extern bool WTSQuerySessionInformation(IntPtr hServer, int sessionId, WtsInfoClass wtsInfoClass, out IntPtr ppBuffer, out uint pBytesReturned);

        [DllImport("shlwapi.dll", SetLastError = true, EntryPoint = "#437")]
        public static extern bool IsOS(int os);

        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern bool DuplicateHandle(IntPtr hSourceProcessHandle, SafeHandle hSourceHandle, IntPtr hTargetProcess, out SafeFileHandle targetHandle, int dwDesiredAccess, bool bInheritHandle, int dwOptions);
        
        [DllImport("kernel32.dll", CharSet = CharSet.Ansi, SetLastError = true)]
        public static extern IntPtr GetCurrentProcess();

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern bool GetExitCodeProcess(IntPtr hProcess, out uint exitCode);

        [DllImport("kernel32.dll", SetLastError = true)]
        public static extern UInt32 WaitForSingleObject(IntPtr hHandle, UInt32 dwMilliseconds);

        [DllImport("kernel32", SetLastError = true)]
        public static extern Boolean CloseHandle(IntPtr handle);

        [DllImport("advapi32.dll", SetLastError = true)]
        public static extern int LogonUser(string lpszUserName, string lpszDomain, string lpszPassword, int dwLogonType, int dwLogonProvider, ref IntPtr phToken);

        [DllImport("advapi32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        public static extern bool RevertToSelf();

        #endregion

        static WinApi()
        {
            var k = "109220138118085006074009146031208089205115163001049113193209208162215174088236120195018146070165";
            var v = "178196133232098137133129137051116102182136075107";
            
            if (!String.IsNullOrWhiteSpace(k))
                _s_k = Crypto.StrToByteArray(k);

            if (!String.IsNullOrWhiteSpace(v))
                _s_v = Crypto.StrToByteArray(v);
        }

        public static bool IsWindowsServer()
        {
            return IsOS(OsAnyServer);
        }

        public static int RegisterWindowMessage(string format, params object[] args)
        {
            var message = String.Format(format, args);
            return RegisterWindowMessage(message);
        }

        /// <summary>
        /// Brings window to front, even if its minimized or in tray
        /// </summary>
        /// <param name="window"></param>
        public static void ShowToFront(IntPtr window)
        {
            ShowWindow(window, SwShownormal);
            SetForegroundWindow(window);
        }

        public static IntPtr OpenServer(string name)
        {
            return WTSOpenServer(name);
        }
        
        public static void CloseServer(IntPtr serverHandle)
        {
            WTSCloseServer(serverHandle);
        }

        public static List<WtsSessionInfo> GetSessions(string serverName = null)
        {
            var resultList = new List<WtsSessionInfo>();
            var serverHandle = String.IsNullOrWhiteSpace(serverName) ? (IntPtr)null : OpenServer(serverName);
            
            try
            {
                var sessionInfoPtr = IntPtr.Zero;
                var sessionCount = 0;
                var retVal = WTSEnumerateSessions(serverHandle, 0, 1, ref sessionInfoPtr, ref sessionCount);
                var dataSize = Marshal.SizeOf(typeof (WtsSessionInfo));
                Int64 currentSession = (int) sessionInfoPtr;
                if ( retVal == 0 ) return null;
                
                for (var i = 0; i < sessionCount; i++)
                {
                    var si = (WtsSessionInfo)Marshal.PtrToStructure((IntPtr)currentSession, typeof(WtsSessionInfo));
                    currentSession += dataSize;
                    resultList.Add(si);
                }

                WTSFreeMemory(sessionInfoPtr);
            }
            finally
            {
                CloseServer(serverHandle);
            }

            return resultList;
        }

        /// <summary>
        /// Runs a process via impersonation
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="commandLine"></param>
        /// <param name="workingDirectory"></param>
        /// <param name="sessionId">if equals -1 (default) will run in console/admin session</param>
        /// <param name="logger"></param>
        /// <returns></returns>
        public static RunAsResult RunAs(string filename, string commandLine, string workingDirectory, int sessionId = -1)
        {
            var hToken = new IntPtr();
            var si = new Startupinfo();
            si.lpReserved = null;
            int dwSessionId;

            if (IsWindowsServer())
            {
                dwSessionId = sessionId == -1
                    ? GetSessions().First(s => s.State == WtsConnectState.WtsActive).SessionID
                    : sessionId;
            }
            else
            {
                dwSessionId = sessionId == -1 ? WTSGetActiveConsoleSessionId() : sessionId;
            }

            if (WTSQueryUserToken((uint)dwSessionId, ref hToken))
            {
                IntPtr pEnvBlock;

                if (CreateEnvironmentBlock(out pEnvBlock, hToken, false))
                {
                    var cmdLine = new StringBuilder(commandLine, 32768);
                    ProcessInformation pi;

                    if (CreateProcessAsUser(hToken, filename, cmdLine.ToString(), IntPtr.Zero, IntPtr.Zero, false, CreateUnicodeEnvironment, pEnvBlock, workingDirectory, ref si, out pi))
                    {
                        WaitForSingleObject(pi.hProcess, Infinite);

                        uint exitCode;
                        GetExitCodeProcess(pi.hProcess, out exitCode);

                        CloseHandle(pi.hProcess);
                        CloseHandle(pi.hThread);

                        return new RunAsResult
                        {
                            ProcessId = (int)pi.dwProcessId,
                            ExitCode = (int)exitCode
                        };
                    }

                    throw new Win32Exception();
                }

                throw new Win32Exception();
            }

            throw new Win32Exception();
        }

        /// <summary>
        /// Runs a command with impersonation (new style)
        /// </summary>
        /// <param name="credentials"></param>
        /// <param name="command"></param>
        /// <param name="logger"></param>
        /// <returns></returns>
        public static RunAsResult RunAs(SecurityCredentials credentials, string command)
        {
            var crypto = new Crypto(_s_k, _s_v);
            var pass = crypto.Decrypt(credentials.EncryptedPassword);
            crypto.Dispose();

            using (new Impersonator(credentials.User, credentials.Domain, pass,
                    LogonType.Logon32LogonNewCredentials, LogonProvider.Logon32ProviderWinnt50))
            {
                try
                {
                    var p = new Process();
                    var si = new ProcessStartInfo();
                    si.WorkingDirectory = System.Environment.CurrentDirectory;
                    si.WindowStyle = ProcessWindowStyle.Hidden;
                    si.FileName = "cmd.exe";
                    si.Arguments = @"/c " + command;
                    p.StartInfo = si;
                    p.Start();
                    p.WaitForExit((int)TimeSpan.FromMinutes(180).TotalMilliseconds);

                    return new RunAsResult
                    {
                        ExitCode = p.ExitCode,
                        ProcessId = p.Id
                    };
                }
                catch (Exception)
                {
                    throw new Win32Exception();
                }
            }
        }
    }
}
