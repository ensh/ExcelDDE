using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using DevExpress.XtraEditors;

using AD.Common.Helpers;
using Core;
using Core.Classes;

namespace GUI.Export
{

    public enum DDEAction
    { 
        Open,
        Close,
        Send,
        Request
    };

    public interface IDDESender
    {
        object Sender { get; }
        void OnComplete(DDEAction action, string CommandText, string ValueText);
    }
    public class ExcelDDEManager
    {
        public static ExcelDDEManager s_instance;

        public static ExcelDDEManager Instance
        {
            get
            {
                return s_instance ?? (s_instance = new ExcelDDEManager());
            }
        }

        public static void FreeInstance() 
        {
            if (s_instance != null)
            {
                s_instance.Dispose();
                s_instance = null;
            }
        }

        uint m_idInst = 0;
        AutoResetEvent m_complete;
        ManualResetEvent m_terminate;
        Thread m_threadForProcess;

	    public ExcelDDEManager()
	    {
		    System = new object();
		    Encoding = Encoding.GetEncoding("windows-1251");

		    m_connections = new ConcurrentDictionary<object, Connection>();
		    m_commands = new ConcurrentQueue<Command>();
		    m_terminate = new ManualResetEvent(false);
		    m_complete = new AutoResetEvent(false);

            // запуск потока экспорта DDE, все системные вызовы DDE должны делаться в контексте одного потока!
            m_threadForProcess = CoreHelper.GetThread(() =>
            {
                try
                {
                    // подключение и все дальнейшие команды dde должны выполнятся в одном и том же потоке!
                    m_DDECallBack = new DDEML.DDECallBackDelegate(DDECallBack);
                    DDEML.DdeInitialize(ref m_idInst, m_DDECallBack, DDEML.APPCMD_CLIENTONLY, 0);

                    do RefreshData(); while (!m_terminate.WaitOne(AppConfig.Common.ExcelExportRefreshDelay));
                }
                catch (Exception e)
                {
                    LogFileManager.AddError("DDE export", e);
                }
            }, true, "Excel DDE processing");
            m_threadForProcess.SetApartmentState(ApartmentState.STA);
            m_threadForProcess.Start();

            //DDECompletitionManager.Instance.Post(() =>
            //    {
            //        m_DDECallBack = new DDEML.DDECallBackDelegate(DDECallBack);
            //        DDEML.DdeInitialize(ref m_idInst, m_DDECallBack, DDEML.APPCLASS_STANDARD, 0);
            //    });
	    }

	    #region IDisposable Members
        public void Dispose()
        {
            if (m_idInst != 0)
            {
                m_terminate.Set();
                m_complete.Dispose();
                m_terminate.Dispose();

                foreach (var conn in m_connections.ToList())
                {
                    Connection c;
                    m_connections.TryRemove(conn.Key, out c);
                    if (conn.Key is IDDESender)
                        ((IDDESender)conn.Key).OnComplete(DDEAction.Close, null, null);
                    if (conn.Key is IDisposable)
                        ((IDisposable)conn.Key).Dispose();
                    DDEML.DdeDisconnect(conn.Value.Conv);
                }

                DDEML.DdeUninitialize(m_idInst);
            }
        }
        #endregion

        public Encoding Encoding { get; set; }
        public object System { get; protected set;}

        DDEML.DDECallBackDelegate m_DDECallBack = null;

        // Обработчик функции обратного вызова
        private IntPtr DDECallBack(uint uType, uint uFmt, IntPtr hConv, IntPtr hsz1, IntPtr hsz2, IntPtr hData, uint dwData1, uint dwData2)
        {
            //LogFileManager.AddInfo("DDE callback uType: " + uType.ToString("X8") + ", hConv: " + (int)hConv + ", dvData1: " + dwData1.ToString("X8"), "call");

            if(uType == DDEML.XTYP_DISCONNECT)
            {
                // сервер закрыт извне
                foreach (var conn in m_connections)
                {
                    if (conn.Value.Conv == hConv)
                    { 
                        // канал обмена данными закрыт извне
                        Connection c;
                        m_connections.TryRemove(conn.Key, out c);

                        if (conn.Key == this.System)  // закрыт системный канал, видимо приложения excel нет вообще, закрываем все
                            FreeInstance();
                        else
                        {
                            // закрыт лист или книга куда осуществлялся экспорт
                            if (conn.Key is IDDESender)
                                ((IDDESender)conn.Key).OnComplete(DDEAction.Close, null, null);
                            if (conn.Key is IDisposable)
                                ((IDisposable)conn.Key).Dispose();
                        }
                    }
                }
            }
            // Все остальные транзакции мы не обрабатываем
            return IntPtr.Zero;
        }

        ConcurrentDictionary<object, Connection> m_connections;

        public bool Open()
        {
            return Open(this.System, "System\0");
        }
        public bool Open(object sender, string topicName)
        {
            Connection conn;
            if (m_connections.TryGetValue(sender, out conn)) return true;

            using (var ds = new DDESender(this, sender) { ValueText = "False"})
            {
                m_commands.Enqueue(new Command(ds, DDEAction.Open, topicName));
                //DDECompletitionManager.Instance.Post(() => ProcessCommand(new Command(ds, DDEAction.Open, topicName)));

                return (WaitComplete()) ? bool.Parse(ds.ValueText) : false;
                //return (m_complete.WaitOne(Timeout.Infinite)) ? bool.Parse(ds.ValueText) : false;
            }
        }
        
        bool InternalOpen(object form, string topicName)
        {
            bool result = false;
            Connection conn;
            if (m_connections.TryGetValue(form, out conn)) return true;

            IntPtr hszService = DDEML.DdeCreateStringHandle(m_idInst, "Excel\0", DDEML.CP_WINANSI);
            IntPtr hszTopic = DDEML.DdeCreateStringHandle(m_idInst, topicName, DDEML.CP_WINANSI);

            // Подключаемся к разделу
            IntPtr hConv = DDEML.DdeConnect(m_idInst, hszService, hszTopic, (IntPtr)null);

            if (hConv != IntPtr.Zero)
            {
                m_connections.TryAdd(form, new Connection(topicName, hConv));
                result = true;
            }

            DDEML.DdeFreeStringHandle(m_idInst, hszTopic);
            DDEML.DdeFreeStringHandle(m_idInst, hszService);
            return result;
        }      
        
        void InternalClose(object form)
        {
            Connection conn;
            if (m_connections.TryGetValue(form, out conn))
            {
                DDEML.DdeDisconnect(conn.Conv);
                m_connections.TryRemove(form, out conn);

                IDDESender sender = form as IDDESender;
                if (sender != null)
                    sender.OnComplete(DDEAction.Close, conn.Name, null);
            }
        }
        public void Close(IDDESender sender)
        {
            Connection conn;
            if (m_connections.TryGetValue(sender, out conn))
            {
                using (var ds = new DDESender(this, sender))
                {
                    m_commands.Enqueue(new Command(ds, DDEAction.Close));
                    //DDECompletitionManager.Instance.Post(() => ProcessCommand(new Command(ds, DDEAction.Close)));
                    WaitComplete();
                    //m_complete.WaitOne(Timeout.Infinite);
                }
            }
        }
        void InternalPokeData(object form, string itemName, string strData)
        {
            Connection conn;
            if (m_connections.TryGetValue(form, out conn))
            {
                IntPtr hszItem = DDEML.DdeCreateStringHandle(m_idInst, itemName, DDEML.CP_WINANSI);
                byte [] data = Encoding.GetBytes(strData);
                IntPtr hszDat = DDEML.DdeCreateDataHandle(m_idInst, data, data.Length, 0, hszItem, DDEML.CF_TEXT, 0);
                uint returnFlags = 0;
                DDEML.DdeClientTransaction(hszDat, -1, conn.Conv, hszItem, DDEML.CF_TEXT, DDEML.XTYP_POKE, Timeout.Infinite, ref returnFlags);
                DDEML.DdeFreeStringHandle(m_idInst, hszItem);
        }
        }
        public void PokeDataAsync(IDDESender sender, string itemName, string strData)
        {
            m_commands.Enqueue(new Command(sender, DDEAction.Send, itemName, strData));
            //DDECompletitionManager.Instance.Post(() => ProcessCommand(new Command(sender, DDEAction.Send, itemName, strData)));
        }
        public void ExecuteMacro(string strText)
        {
            using (var ds = new DDESender(this, this.System))
            {
                m_commands.Enqueue(new Command(ds, DDEAction.Send, strText));
                //DDECompletitionManager.Instance.Post(() => ProcessCommand(new Command(ds, DDEAction.Send, strText)));
                WaitComplete();
            }
        }

        bool WaitComplete()
        {
            bool result = false;
            while (!(result = m_complete.WaitOne(AppConfig.Common.WaitExcelTimeout)))
            {
                if (XtraMessageBox.Show(MainForm.Instance.activeControl,
                    "Истекло время ожидания ответа от сервера Excel.\r\nПродолжить ожидание?", "ВНИМАНИЕ",
                    MessageBoxButtons.OKCancel, MessageBoxIcon.Warning)
                    != DialogResult.OK) return false;
            }
            return result;
        }

        public void ExecuteMacroAsync(IDDESender sender, string strText)
        {
            m_commands.Enqueue(new Command(sender, DDEAction.Send, strText));
            //DDECompletitionManager.Instance.Post(() => ProcessCommand(new Command(sender, DDEAction.Send, strText)));
        }
        void InternalExecuteMacro(object form, string strText)
        {
            Connection conn;
            if (m_connections.TryGetValue(form, out conn))
            {
                uint idErr = 0;
                byte[] data = Encoding.GetBytes(strText);
                IntPtr hszDat = DDEML.DdeCreateDataHandle(m_idInst, data, data.Length, 0, IntPtr.Zero, DDEML.CF_TEXT, 0);
                if (hszDat == IntPtr.Zero)
                {
                    idErr = DDEML.DdeGetLastError(m_idInst);
                }
                uint returnFlags = 0;
                DDEML.DdeClientTransaction(hszDat, -1, conn.Conv, IntPtr.Zero, DDEML.CF_TEXT, DDEML.XTYP_EXECUTE, Timeout.Infinite, ref returnFlags);
            }
        }
        public void Request(string itemName, out string strData)
        {
            Request(this.System, itemName, out strData);
        }
        public void Request(object sender, string itemName, out string strData)
        {
            using (var ds = new DDESender(this, sender))
            {
                m_commands.Enqueue(new Command(ds, DDEAction.Request, itemName));
                //DDECompletitionManager.Instance.Post(() => ProcessCommand(new Command(ds, DDEAction.Request, itemName)));
                strData = (WaitComplete()) ? ds.ValueText : null;
            }
        }
        void InternalRequest(object form, string itemName, out string strData)
        {
            byte[] data;
            InternalRequest(form, itemName, out data);
            if (data.Length == 0) strData = String.Empty;
            strData = Encoding.GetString(data);
        }
        void InternalRequest(string itemName, out byte[] data)
        {
            InternalRequest(this.System, itemName, out data);
        }
        public void InternalRequest(object form, string itemName, out byte[] data)
        {
            Connection conn;
            data = new byte[0];
            if (m_connections.TryGetValue(form, out conn))
            {
                IntPtr hszItem = DDEML.DdeCreateStringHandle(m_idInst, itemName, DDEML.CP_WINANSI);
                uint returnFlags = 0;
                IntPtr dataHandle = DDEML.DdeClientTransaction(IntPtr.Zero, 0, conn.Conv, hszItem, DDEML.CF_TEXT, DDEML.XTYP_REQUEST, 1000, ref returnFlags);
                DDEML.DdeFreeStringHandle(m_idInst, hszItem);

                if (dataHandle != IntPtr.Zero)
                {
                    uint length = DDEML.DdeGetData(dataHandle, null, 0, 0);
                    data = new byte[length];
                    length = DDEML.DdeGetData(dataHandle, data, (uint)data.Length, 0);

                    // Free the data handle created by the server.
                    DDEML.DdeFreeDataHandle(dataHandle);
                }
            }            
        }

        ConcurrentQueue<Command> m_commands;
        void RefreshData()
        {
            bool result;
            Command command;
            while (m_commands.TryDequeue(out command))
            {
                switch (command.Action)
                {
                    case DDEAction.Open:
                        result = InternalOpen(command.Sender.Sender, command.CommandText);
                        command.Sender.OnComplete(DDEAction.Open, command.CommandText, result.ToString());
                        break;
                    case DDEAction.Close:
                        InternalClose(command.Sender.Sender);
                        command.Sender.OnComplete(DDEAction.Close, null, null);
                        break;
                    case DDEAction.Request:
                        string request = null;
                        InternalRequest(command.Sender.Sender, command.CommandText, out request);
                        command.Sender.OnComplete(DDEAction.Request, command.CommandText, request);
                        break;
                    case DDEAction.Send:
                        if (String.IsNullOrEmpty(command.ValueText))
                        {
                            InternalExecuteMacro(command.Sender.Sender, command.CommandText);
                            command.Sender.OnComplete(DDEAction.Send, command.CommandText, null);
                        }
                        else
                        {
                            InternalPokeData(command.Sender.Sender, command.CommandText, command.ValueText);
                            command.Sender.OnComplete(DDEAction.Send, command.CommandText, command.ValueText);
                        }
                        break;
                }
            }
        }

        void ProcessCommand(Command command)
        {
            bool result;

            switch (command.Action)
            {
                case DDEAction.Open:
                    result = InternalOpen(command.Sender.Sender, command.CommandText);
                    command.Sender.OnComplete(DDEAction.Open, command.CommandText, result.ToString());
                    break;
                case DDEAction.Close:
                    InternalClose(command.Sender.Sender);
                    command.Sender.OnComplete(DDEAction.Close, null, null);
                    break;
                case DDEAction.Request:
                    string request = null;
                    InternalRequest(command.Sender.Sender, command.CommandText, out request);
                    command.Sender.OnComplete(DDEAction.Request, command.CommandText, request);
                    break;
                case DDEAction.Send:
                    if (String.IsNullOrEmpty(command.ValueText))
                    {
                        InternalExecuteMacro(command.Sender.Sender, command.CommandText);
                        command.Sender.OnComplete(DDEAction.Send, command.CommandText, null);
                    }
                    else
                    {
                        InternalPokeData(command.Sender.Sender, command.CommandText, command.ValueText);
                        command.Sender.OnComplete(DDEAction.Send, command.CommandText, command.ValueText);
                    }
                    break;
            }
        }

        struct Connection
        {
            public readonly string Name;
            public readonly IntPtr Conv;

            public Connection(string name, IntPtr conv)
            {
                Conv = conv;
                Name = name;
            }
        }

        struct Command
        {
            public readonly IDDESender Sender;
            public readonly DDEAction Action;
            public readonly string CommandText;
            public readonly string ValueText;

            public Command(IDDESender sender, DDEAction action, string commandText = null, string valueText = null)
            {
                Sender = sender;
                Action = action;
                CommandText = commandText;
                ValueText = valueText;
            }
        }

        class DDESender : IDisposable, IDDESender
        {
            ExcelDDEManager _owner;
            public DDESender(ExcelDDEManager manager, object sender)
            {
                _owner = manager;
                Sender = sender;
                Monitor.Enter(_owner.System);
                _owner.m_complete.Reset();
            }
            #region IDDESender Members
            public object Sender { get; protected set; }

            public void OnComplete(DDEAction action, string commandText, string valueText)
            {
                CommandText = commandText;
                ValueText = valueText;
                _owner.m_complete.Set();
            }

            public string CommandText;
            public string ValueText;
            #endregion

            #region IDisposable Members

            public void Dispose()
            {
                Monitor.Exit(_owner.System);
            }

            #endregion
        }
    }

    public sealed class DDECompletitionManager : Core.QueuedCompletitionManager
    {
        static DDECompletitionManager _instance;
        public static DDECompletitionManager Instance
        {
            get
            {
                return _instance ?? (_instance = new DDECompletitionManager());
            }
        }
        public DDECompletitionManager(): base()
        {
            _actionQueue = new ConcurrentQueue<DDEAction>();
            Start();
        }

        protected override void Apply()
        {
            DDEAction exaction;
            if (_actionQueue.TryDequeue(out exaction))
            {
                exaction.@do();
            }
        }
        public void Post(Action a)
        {
            _actionQueue.Enqueue(new DDEAction(a));
            base.Post();
        }
        public struct DDEAction
        {
            Action action;
            public DDEAction(Action a)
            {
                action = a;
            }
            public static implicit operator Action(DDEAction a)
            {
                return a.action;
            }
            public void @do()
            {
                action();
            }
        }

        ConcurrentQueue<DDEAction> _actionQueue;
    }
}
