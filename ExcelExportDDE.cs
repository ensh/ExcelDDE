using System;
using System.Collections.Generic;
using System.Collections.Concurrent;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;

using DevExpress.XtraEditors;
using DevExpress.XtraGrid.Views.BandedGrid;
using DevExpress.XtraGrid.Views.BandedGrid.Handler;
using DevExpress.XtraGrid.Views.BandedGrid.ViewInfo;

using AD.Common.Helpers;
using Core;

namespace GUI.Export
{
    public struct ExcelDDEExportSettings
    {
        public string BookName;
        public string SheetName;
        public int ColStart;
        public int RowStart;
        public bool AddTitles;
        public bool AutoStart;
        public bool OnlyVisible;

        public ExcelDDEExportSettings(params object[] settings)
        { 
            BookName = (settings.Length > 0) ? (string)settings[0] : null;
            SheetName = (settings.Length > 1) ? (string)settings[1] : null;
            ColStart = (settings.Length > 2) ? (int)settings[2] : 1;
            RowStart = (settings.Length > 3) ? (int)settings[3] : 1;
            AddTitles = (settings.Length > 4) ? (bool)settings[4] : false;
            OnlyVisible = (settings.Length > 5) ? (bool)settings[5] : true;
            AutoStart = (settings.Length > 6) ? (bool)settings[6] : false;
        }

        public static implicit operator List<string>(ExcelDDEExportSettings s)
        {
            return new List<string>()
            {
                s.BookName, s.SheetName, s.ColStart.ToString(), s.RowStart.ToString(), s.AddTitles.ToString(),
                s.OnlyVisible.ToString(), s.AutoStart.ToString(),
            };
        }

        public static implicit operator ExcelDDEExportSettings(List<string> s)
        {
            object [] settings = new object[8];

            if (s == null) s = new List<string>(0);

            settings[0] = (s.Count > 0) ? s[0] : null;
            settings[1] = (s.Count > 1) ? s[1] : null;
            int i;
            settings[2] = (s.Count > 2 && int.TryParse(s[2], out i)) ? i : 1;
            settings[3] = (s.Count > 3 && int.TryParse(s[3], out i)) ? i : 1;
            bool b;
            settings[4] = (s.Count > 4 && bool.TryParse(s[4], out b)) ? b : false;
            settings[5] = (s.Count > 5 && bool.TryParse(s[5], out b)) ? b : false;
            settings[6] = (s.Count > 6 && bool.TryParse(s[6], out b)) ? b : false;

            return new ExcelDDEExportSettings(settings);
        }
    }

    public class ExcelExportDDE : IDisposable, IDDESender
    {
        IExportableView m_view;
        IEnumerable<BandedGridColumn> m_columns;
        ExcelDDEExportSettings m_settings;

        public ExcelExportDDE(IExportableView view, params object[] settings)
            : this(view, new ExcelDDEExportSettings(settings))
        {
        }
        public ExcelExportDDE(IExportableView view, ExcelDDEExportSettings settings)
        {
            m_view = view;
            m_settings = settings;
        }

        public void Open()
        {
            if (Active = InternalOpen())
            {
                m_colMax = m_colMin = m_settings.ColStart;
                m_columns = m_view.ExportColumns(m_settings.OnlyVisible).ToList();
                m_colMax += m_columns.Count() -1;

                // блокировка изменений таблицы на время первоначальной выгрузки
                using (var viewLocker = new ViewLocker(m_view))
                {
                    AllExport(this, null);
                    m_view.DatasetChanged += ResetExport;
                    m_view.DatasetChanged += AllExport;
                }
            }
        }
        
        public event Action OnDisposed;

        public bool Active
        {
            get;  protected set;
        }

        public ExcelDDEExportSettings ExportSettings
        {
            get
            {
                return m_settings;
            }
        }

        bool InternalOpen()
        {
            string bookname;
            string topicList = null;

            string fileName = (String.IsNullOrEmpty(System.IO.Path.GetDirectoryName(m_settings.BookName))) ?
                Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\\" + m_settings.BookName : m_settings.BookName;

            // открываем уже существующий документ и лист
            if (!ExcelDDEManager.Instance.Open(this, Command.DDEName(System.IO.Path.GetFileName(fileName), m_settings.SheetName)))
            {
                // что-то не создано... открываем системный dde
                if(!ExcelDDEManager.Instance.Open())
                {
                    // нужно запустить сам excel
                    System.Diagnostics.Process p = System.Diagnostics.Process.Start("excel.exe");
                    p.WaitForInputIdle();

                    // ну вообще с dde траблы
                    if (!ExcelDDEManager.Instance.Open()) 
                        return false;
                }

                // открыть имеющийся файл
                if (System.IO.File.Exists(fileName))
                {
                    ExcelDDEManager.Instance.ExecuteMacro(String.Concat("[OPEN(\"", fileName, "\")]\0"));

                    int wait = 2;
                    while(wait > 0)
                    {
                        // попытаемся переоткрыть лист
                        if (ExcelDDEManager.Instance.Open(this, Command.DDEName(System.IO.Path.GetFileName(fileName), m_settings.SheetName)))
                        {
                            m_settings.BookName = fileName;
                            ExcelDDEManager.Instance.ExecuteMacroAsync(this, String.Concat("[WORKBOOK.ACTIVATE(\"", m_settings.SheetName, "\")]\0"));
                            return true;
                        }
                        Thread.Sleep(1000);// дадим книге загрузится...
                        wait--;
                    }

                    // создаем новый лист
                    ExcelDDEManager.Instance.ExecuteMacro("[WORKBOOK.INSERT(1)]\0");
                }
                else
                {
                    // новая книга
                    ExcelDDEManager.Instance.ExecuteMacro("[NEW(1)]\0");
                    // переименовываем
                    ExcelDDEManager.Instance.ExecuteMacro(String.Concat("[SAVE.AS(\"", fileName, "\")]\0"));
                }

                // находим топики загруженной книги
                topicList = null;
                Stopwatch startWaiting = Stopwatch.StartNew();
                while (String.IsNullOrEmpty(topicList))
                {
                    // если excel выкинул диалог, список пустой или нулл
                    ExcelDDEManager.Instance.Request("Topics\0", out topicList);

                    if (startWaiting.ElapsedMilliseconds > AppConfig.Common.WaitExcelTimeout)
                    {
                        // что-то пошло не так...
                        if (XtraMessageBox.Show(MainForm.Instance.activeControl,
                            "Истекло время ожидания ответа от сервера Excel.\r\nПродолжить ожидание?", "ВНИМАНИЕ",
                            MessageBoxButtons.OKCancel, MessageBoxIcon.Warning)
                            != DialogResult.OK) return false;

                        startWaiting = Stopwatch.StartNew();
                    }
                }

                //if (topicList != null)
                {
                    bookname = System.IO.Path.GetFileName(fileName);
                    string[] topics = topicList.Split('\t')  // нужны топики только выбранной книги: [книга]лист
                        .Where(t => t.Contains(bookname))
                        .ToArray();

                    // переименовываем последний лист в списке
                    if (topics.Length > 0) // в книге хотя бы один лист
                    {
                        string[] tname = topics[topics.Length - 1].Split(']');
                        if (tname.Length == 2) // в кни
                        {
                            bookname = tname[0].Substring(1); // имя книги может не совпадать с фактическим
                            // переименовываем последний лист в списке
                            ExcelDDEManager.Instance.ExecuteMacro(String.Concat("[WORKBOOK.NAME(\"", tname[1], "\",\"", m_settings.SheetName, "\")]\0"));
                        }
                    }
                }

                // попытаемся переоткрыть лист
                if (ExcelDDEManager.Instance.Open(this, Command.DDEName(bookname, m_settings.SheetName)))
                {
                    m_settings.BookName = System.IO.Path.GetDirectoryName(fileName) + "\\" + bookname; 
                    ExcelDDEManager.Instance.ExecuteMacroAsync(this, String.Concat("[WORKBOOK.ACTIVATE(\"", m_settings.SheetName, "\")]\0"));
                    return true;
                }
                return false;
            }
            return true;
        }

        int m_colMax, m_rowMax, m_rowMin, m_colMin;

        void ExportTitles()
        {
            ExcelDDEManager.Instance.PokeDataAsync(this, Command.Range(m_rowMin++, m_colMin, m_colMax), 
                m_view.ColumnTitles("\t", m_columns) + '\0');

            ExcelDDEManager.Instance.PokeDataAsync(this, Command.Range(m_rowMin++, m_colMin, m_colMax),
                m_view.ColumnCaptions("\t", m_columns) + '\0');

            m_rowMax = m_rowMin;
        }

        void ExportValues()
        {
            if (m_view.RowCount == 0) return;

            StringBuilder sb = new StringBuilder();
            int rowCount = 0;
            foreach (var s in m_view.TableValues("\t", m_columns))
            {
                sb.Append(s);
                rowCount++;
            }

            if (rowCount > 0)
            {
                m_rowMax += rowCount;
                ExcelDDEManager.Instance.PokeDataAsync(this, Command.Range(m_rowMin, m_rowMax, m_colMin, m_colMax), sb.Append('\0').ToString());
            }
        }

        void AllExport(object sender, EventArgs e)
        {
            int rowNumber = m_settings.RowStart;
            StringBuilder sb = new StringBuilder();

            if (m_settings.AddTitles)
            {
                sb.AppendLine(m_view.ColumnTitles("\t", m_columns));
                rowNumber++;
                sb.AppendLine(m_view.ColumnCaptions("\t", m_columns));
                rowNumber++;
            }

            foreach (var s in m_view.TableValues("\t", m_columns))
            {
                sb.Append(s);
                rowNumber++;
            }

            if (sb.Length > 0)
            {
                ExcelDDEManager.Instance.PokeDataAsync(this,
                    Command.Range(m_settings.RowStart, (m_rowMax = rowNumber) -1, m_colMin, m_colMax), sb.Append('\0').ToString());
            }
        }

        void ResetExport(object sender, EventArgs e)
        {
            StringBuilder sb = new StringBuilder();

            var cols = m_columns.Select((c) => "").ToArray();
            for (int i = m_settings.RowStart; i < m_rowMax; i++)
            {
                var s =  String.Join("\t", cols);
                sb.AppendLine(s);
            }

            if (sb.Length > 0)
            {
                ExcelDDEManager.Instance.PokeDataAsync(this,
                    Command.Range(m_settings.RowStart, m_rowMax-1, m_colMin, m_colMax), sb.Append('\0').ToString());
            }
        }

        #region IDisposable Members
        public void Dispose()
        {
            using (var viewLocker = new ViewLocker(m_view))
            {
                m_view.DatasetChanged -= ResetExport;
                m_view.DatasetChanged -= AllExport;
            }

            m_view = null;
            m_columns = null;

            ExcelDDEManager.Instance.Close(this);

            Active = false;

            if (OnDisposed != null)
                OnDisposed();
        }
        #endregion

        static class Command
        {
            public static string Insert(int row, int colStart, int colFinish)
            { 
                return String.Concat("[SELECT(\"", Range(row, colStart, colFinish, ""), "\")][INSERT(2)][SELECT(\"\")]\0");
            }
            public static string Insert(int rowStart, int RowFinish, int colStart, int colFinish)
            {
                return String.Concat("[SELECT(\"", Range(rowStart, RowFinish, colStart, colFinish, ""), "\")][INSERT(2)][SELECT(\"\")]\0");
            }
            public static string Delete(int row, int colStart, int colFinish)
            {
                return String.Concat("[SELECT(\"", Range(row, colStart, colFinish, ""), "\")][EDIT.DELETE(2)][SELECT(\"\")]\0");
            }
            public static string Delete(int rowStart, int RowFinish, int colStart, int colFinish)
            {
                return String.Concat("[SELECT(\"", Range(rowStart, RowFinish, colStart, colFinish, ""), "\")][EDIT.DELETE(2)][SELECT(\"\")]\0");
            }

            public static string Range(int row, int colStart, int colFinish, string lastChar = "\0")
            {
                string R = "R" + row.ToString();
                string C1 = "C" + colStart.ToString();
                string C2 = "C" + colFinish.ToString();
                return String.Concat(R, C1, ":", R, C2, lastChar);
            }
            public static string Range(int rowStart, int RowFinish, int colStart, int colFinish, string lastChar = "\0")
            {
                string R1 = "R" + rowStart.ToString();
                string R2 = "R" + RowFinish.ToString();
                string C1 = "C" + colStart.ToString();
                string C2 = "C" + colFinish.ToString();
                return String.Concat(R1, C1, ":", R2, C2, lastChar);
            }
            public static string DDEName(string bookName, string sheetName)
            {
                return String.Concat("[", bookName, "]", sheetName ,"\0");
            }
        }

        #region IDDESender Members

        public object Sender { get { return this; } }

        public void OnComplete(DDEAction action, string CommandText, string ValueText)
        {
            switch (action)
            { 
                case DDEAction.Open:
                    //LogFileManager.AddInfo("Open: " + CommandText ?? "" + "=>" + ValueText??"" , "export");
                    break;
                case DDEAction.Close:
                    //LogFileManager.AddInfo("Close: " + CommandText ?? "", "export");
                    Active = false;
                    break;
                case DDEAction.Send:
                    //LogFileManager.AddInfo("Command: " + CommandText??"", "export");
                    //if (!String.IsNullOrEmpty(ValueText))
                    //    LogFileManager.AddInfo("Value: " + ValueText, "export");
                    break;
            }
        }

        #endregion
    }
}
