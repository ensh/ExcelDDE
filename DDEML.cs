using System;
using System.Runtime.InteropServices;

namespace GUI.Export
{
    public static class DDEML
    {
        public const int MAX_STRING_SIZE = 255;

        public const int APPCMD_CLIENTONLY = unchecked((int)0x00000010);
        public const int APPCMD_FILTERINITS = unchecked((int)0x00000020);
        public const int APPCMD_MASK = unchecked((int)0x00000FF0);
        public const int APPCLASS_STANDARD = unchecked((int)0x00000000);
        public const int APPCLASS_MONITOR = unchecked((int)0x00000001);
        public const int APPCLASS_MASK = unchecked((int)0x0000000F);

        public const int CBR_BLOCK = unchecked((int)0xFFFFFFFF);

        public const int CBF_FAIL_SELFCONNECTIONS = unchecked((int)0x00001000);
        public const int CBF_FAIL_CONNECTIONS = unchecked((int)0x00002000);
        public const int CBF_FAIL_ADVISES = unchecked((int)0x00004000);
        public const int CBF_FAIL_EXECUTES = unchecked((int)0x00008000);
        public const int CBF_FAIL_POKES = unchecked((int)0x00010000);
        public const int CBF_FAIL_REQUESTS = unchecked((int)0x00020000);
        public const int CBF_FAIL_ALLSVRXACTIONS = unchecked((int)0x0003f000);
        public const int CBF_SKIP_CONNECT_CONFIRMS = unchecked((int)0x00040000);
        public const int CBF_SKIP_REGISTRATIONS = unchecked((int)0x00080000);
        public const int CBF_SKIP_UNREGISTRATIONS = unchecked((int)0x00100000);
        public const int CBF_SKIP_DISCONNECTS = unchecked((int)0x00200000);
        public const int CBF_SKIP_ALLNOTIFICATIONS = unchecked((int)0x003c0000);

        public const int CF_TEXT = 1;

        public const int CP_WINANSI = 1004;
        public const int CP_WINUNICODE = 1200;

        public const int DDE_FACK = unchecked((int)0x8000);
        public const int DDE_FBUSY = unchecked((int)0x4000);
        public const int DDE_FDEFERUPD = unchecked((int)0x4000);
        public const int DDE_FACKREQ = unchecked((int)0x8000);
        public const int DDE_FRELEASE = unchecked((int)0x2000);
        public const int DDE_FREQUESTED = unchecked((int)0x1000);
        public const int DDE_FAPPSTATUS = unchecked((int)0x00ff);
        public const int DDE_FNOTPROCESSED = unchecked((int)0x0000);

        public const int DMLERR_NO_ERROR = unchecked((int)0x0000);
        public const int DMLERR_FIRST = unchecked((int)0x4000);
        public const int DMLERR_ADVACKTIMEOUT = unchecked((int)0x4000);
        public const int DMLERR_BUSY = unchecked((int)0x4001);
        public const int DMLERR_DATAACKTIMEOUT = unchecked((int)0x4002);
        public const int DMLERR_DLL_NOT_INITIALIZED = unchecked((int)0x4003);
        public const int DMLERR_DLL_USAGE = unchecked((int)0x4004);
        public const int DMLERR_EXECACKTIMEOUT = unchecked((int)0x4005);
        public const int DMLERR_INVALIDPARAMETER = unchecked((int)0x4006);
        public const int DMLERR_LOW_MEMORY = unchecked((int)0x4007);
        public const int DMLERR_MEMORY_ERROR = unchecked((int)0x4008);
        public const int DMLERR_NOTPROCESSED = unchecked((int)0x4009);
        public const int DMLERR_NO_CONV_ESTABLISHED = unchecked((int)0x400A);
        public const int DMLERR_POKEACKTIMEOUT = unchecked((int)0x400B);
        public const int DMLERR_POSTMSG_FAILED = unchecked((int)0x400C);
        public const int DMLERR_REENTRANCY = unchecked((int)0x400D);
        public const int DMLERR_SERVER_DIED = unchecked((int)0x400E);
        public const int DMLERR_SYS_ERROR = unchecked((int)0x400F);
        public const int DMLERR_UNADVACKTIMEOUT = unchecked((int)0x4010);
        public const int DMLERR_UNFOUND_QUEUE_ID = unchecked((int)0x4011);
        public const int DMLERR_LAST = unchecked((int)0x4011);

        public const int DNS_REGISTER = unchecked((int)0x0001);
        public const int DNS_UNREGISTER = unchecked((int)0x0002);
        public const int DNS_FILTERON = unchecked((int)0x0004);
        public const int DNS_FILTEROFF = unchecked((int)0x0008);

        public const int EC_ENABLEALL = unchecked((int)0x0000);
        public const int EC_ENABLEONE = unchecked((int)0x0080);
        public const int EC_DISABLE = unchecked((int)0x0008);
        public const int EC_QUERYWAITING = unchecked((int)0x0002);

        public const int HDATA_APPOWNED = unchecked((int)0x0001);

        public const int MF_HSZ_INFO = unchecked((int)0x01000000);
        public const int MF_SENDMSGS = unchecked((int)0x02000000);
        public const int MF_POSTMSGS = unchecked((int)0x04000000);
        public const int MF_CALLBACKS = unchecked((int)0x08000000);
        public const int MF_ERRORS = unchecked((int)0x10000000);
        public const int MF_LINKS = unchecked((int)0x20000000);
        public const int MF_CONV = unchecked((int)0x40000000);

        public const int MH_CREATE = 1;
        public const int MH_KEEP = 2;
        public const int MH_DELETE = 3;
        public const int MH_CLEANUP = 4;

        public const int QID_SYNC = unchecked((int)0xFFFFFFFF);
        public const int TIMEOUT_ASYNC = unchecked((int)0xFFFFFFFF);

        public const int XTYPF_NOBLOCK = unchecked((int)0x0002);
        public const int XTYPF_NODATA = unchecked((int)0x0004);
        public const int XTYPF_ACKREQ = unchecked((int)0x0008);
        public const int XCLASS_MASK = unchecked((int)0xFC00);
        public const int XCLASS_BOOL = unchecked((int)0x1000);
        public const int XCLASS_DATA = unchecked((int)0x2000);
        public const int XCLASS_FLAGS = unchecked((int)0x4000);
        public const int XCLASS_NOTIFICATION = unchecked((int)0x8000);
        public const int XTYP_ERROR = unchecked((int)(0x0000 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK));
        public const int XTYP_ADVDATA = unchecked((int)(0x0010 | XCLASS_FLAGS));
        public const int XTYP_ADVREQ = unchecked((int)(0x0020 | XCLASS_DATA | XTYPF_NOBLOCK));
        public const int XTYP_ADVSTART = unchecked((int)(0x0030 | XCLASS_BOOL));
        public const int XTYP_ADVSTOP = unchecked((int)(0x0040 | XCLASS_NOTIFICATION));
        public const int XTYP_EXECUTE = unchecked((int)(0x0050 | XCLASS_FLAGS));
        public const int XTYP_CONNECT = unchecked((int)(0x0060 | XCLASS_BOOL | XTYPF_NOBLOCK));
        public const int XTYP_CONNECT_CONFIRM = unchecked((int)(0x0070 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK));
        public const int XTYP_XACT_COMPLETE = unchecked((int)(0x0080 | XCLASS_NOTIFICATION));
        public const int XTYP_POKE = unchecked((int)(0x0090 | XCLASS_FLAGS));
        public const int XTYP_REGISTER = unchecked((int)(0x00A0 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK));
        public const int XTYP_REQUEST = unchecked((int)(0x00B0 | XCLASS_DATA));
        public const int XTYP_DISCONNECT = unchecked((int)(0x00C0 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK));
        public const int XTYP_UNREGISTER = unchecked((int)(0x00D0 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK));
        public const int XTYP_WILDCONNECT = unchecked((int)(0x00E0 | XCLASS_DATA | XTYPF_NOBLOCK));
        public const int XTYP_MONITOR = unchecked((int)(0x00F0 | XCLASS_NOTIFICATION | XTYPF_NOBLOCK));
        public const int XTYP_MASK = unchecked((int)0x00F0);
        public const int XTYP_SHIFT = unchecked((int)0x0004);

        [DllImport("user32.dll", EntryPoint = "DdeInitialize", CharSet = CharSet.Ansi)]
        internal static extern uint DdeInitialize(ref uint pidInst, DDECallBackDelegate pfnCallback, uint afCmd, uint ulRes);

        internal delegate IntPtr DDECallBackDelegate(
            uint wType, //Код транзакции
            uint wFmt, // Формат данных
            IntPtr hConv, // Идентификатор канала
            IntPtr hsz1, // Идентификатор строки (topic)
            IntPtr hsz2,  // Идентификатор строки (item)
            IntPtr hData, // Идентификатор глобальной области данных, где находятся данные
            uint dwData1, // Идентификатор транзакции
            uint dwData2 // Дополнительный статус операции
         );

        [DllImport("user32.dll", EntryPoint = "DdeUninitialize", CharSet = CharSet.Ansi)]
        internal static extern bool DdeUninitialize(uint idInst);

        [DllImport("user32.dll", EntryPoint = "DdeCreateDataHandle", CharSet = CharSet.Ansi)]
        public static extern IntPtr DdeCreateDataHandle(uint idInst, byte[] pSrc, int cb, uint cbOff, IntPtr hszItem, uint wFmt, int afCmd);

        [DllImport("user32.dll", EntryPoint = "DdeCreateStringHandle", CharSet = CharSet.Ansi)]
        internal static extern IntPtr DdeCreateStringHandle(uint idInst, string psz, int iCodePage);

        [DllImport("user32.dll", EntryPoint = "DdeFreeDataHandle", CharSet = CharSet.Ansi)]
        public static extern bool DdeFreeDataHandle(IntPtr hData);

        [DllImport("user32.dll", EntryPoint = "DdeFreeStringHandle", CharSet = CharSet.Ansi)]
        internal static extern bool DdeFreeStringHandle(uint idInst, IntPtr hsz);

        [DllImport("user32.dll", EntryPoint = "DdeConnect", CharSet = CharSet.Ansi)]
        internal static extern IntPtr DdeConnect(uint idInst, IntPtr hszService, IntPtr hszTopic, IntPtr pCC);

        [DllImport("user32.dll", EntryPoint = "DdeDisconnect", CharSet = CharSet.Ansi)]
        internal static extern bool DdeDisconnect(IntPtr hConv);

        [DllImport("user32.dll", EntryPoint = "DdeClientTransaction", CharSet = CharSet.Ansi)]
        internal static extern IntPtr DdeClientTransaction(IntPtr pData, int cbData, IntPtr hConv, IntPtr hszItem, uint uFmt, uint uType, int dwTimeout, ref uint pdwResult);

        [DllImport("user32.dll", EntryPoint = "DdeGetData", CharSet = CharSet.Ansi)]
        internal static extern uint DdeGetData(IntPtr hData, [Out] byte[] pDst, uint cbMax, uint cbOff);

        [DllImport("user32.dll", EntryPoint = "DdeGetData", CharSet = CharSet.Ansi)]
        internal static extern uint DdeAddData(IntPtr hData, [In] byte[] pSrc, uint cbMax, uint cbOff);

        [DllImport("user32.dll", EntryPoint = "DdeQueryString", CharSet = CharSet.Ansi)]
        internal static extern uint DdeQueryString(uint idInst, IntPtr hsz, string psz, uint cchMax, int iCodePage);

        [DllImport("user32.dll", EntryPoint = "DdeGetLastError", CharSet = CharSet.Ansi)]
        public static extern uint DdeGetLastError([In] uint idInst);

    }

}
