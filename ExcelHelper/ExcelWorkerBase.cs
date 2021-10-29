using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using CustomTools;

namespace ExcelHelper
{
    /// <summary>
    /// Базовый класс для работы с Microsoft Excel
    /// </summary>
    public abstract class ExcelWorkerBase : IDisposable
    {
        protected Application _excel;
        protected Workbook _workbook;
        protected Worksheet _worksheet;
        protected string _path;

        public event System.Action Closing;
        public event System.Action Closed;
        public event System.Action Opened;
        public event System.Action Loaded;
        public event System.Action SheetChanged;

        /// <summary>
        /// Открывает файл по заданному пути
        /// </summary>
        /// <param name="path">Путь файла</param>
        public ExcelWorkerBase(string path)
        {
            try
            {
                _excel = new Application();
                _workbook = _excel.Workbooks.Open(path);

                Loaded?.Invoke();
            }
            catch (Exception e)
            {
                Logger.Log(e);
                Close();
            }
        }
        /// <summary>
        /// Запускает приложение
        /// </summary>
        public ExcelWorkerBase()
        {
            _excel = new Application();

            Loaded?.Invoke();
        }

        /// <summary>
        /// Сохраняет изменения в файле
        /// </summary>
        public void Save()
        {
            _workbook.Save();
        }
        /// <summary>
        /// Сохраняет файл по заданному пути
        /// </summary>
        /// <param name="path">Путь сохранения</param>
        public void SaveAs(string path)
        {
            _workbook.SaveAs(path);
        }

        /// <summary>
        /// Создает новый файл без сохранения
        /// </summary>
        public void NewFile()
        {
            _workbook = _excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            _worksheet = _workbook.Worksheets[1];
        }
        /// <summary>
        /// Создает новый файл по заданному пути
        /// </summary>
        /// <param name="path">Путь сохранения</param>
        public void NewFile(string path)
        {
            _workbook = _excel.Workbooks.Add(XlWBATemplate.xlWBATWorksheet);
            _worksheet = _workbook.Worksheets[1];
            SaveAs(path);
        }
        /// <summary>
        /// Создает новую вкладку в документе на первом месте
        /// </summary>
        public void NewSheet()
        {
            //Worksheet tmp = _workbook.Worksheets.Add(After: _workbook.Worksheets[_workbook.Worksheets.Count]);
            Worksheet tmp = _workbook.Worksheets.Add();
        }
        /// <summary>
        /// Создает новуу вкладку с заданным именем
        /// </summary>
        /// <param name="name">Имя вкладки</param>
        public void NewSheet(string name)
        {
            //Worksheet tmp = _workbook.Worksheets.Add(After: _workbook.Worksheets[_workbook.Worksheets.Count]);
            Worksheet tmp = _workbook.Worksheets.Add();
            tmp.Name = name;
        }
        /// <summary>
        /// Удаляет вкладку с выбранным номером
        /// </summary>
        /// <param name="sheet">Номер вкладки</param>
        public void DeleteSheet(int sheet)
        {
            _workbook.Worksheets[sheet].Delete();
        }
        public void DeleteSheet(string name)
        {
            _workbook.Sheets[name].Delete();
        }
        /// <summary>
        /// Переводит фокус на вкладку с нужным именем
        /// </summary>
        /// <param name="name">Имя вкладки</param>
        public void SelectSheet(string name)
        {
            _worksheet = _workbook.Sheets[name];
            _worksheet.Activate();

            SheetChanged?.Invoke();
        }
        /// <summary>
        /// Открывает вкладку определенного номера
        /// </summary>
        /// <param name="sheet">Номер вкладки (нумерация с 0)</param>
        public void SelectSheet(int sheet)
        {
            _worksheet = _workbook.Sheets[++sheet];
            _worksheet.Activate();

            SheetChanged?.Invoke();
        }

        /// <summary>
        /// Открывает файл по его пути
        /// </summary>
        /// <param name="path">Путь файла</param>
        public void Open(string path)
        {
            _workbook = _excel.Workbooks.Open(path);
            SelectSheet(0);

            Opened?.Invoke();
        }
        /// <summary>
        /// Закрывает файл с сохранением изменений
        /// </summary>
        public void Close()
        {
            Closing?.Invoke();

            _workbook?.Save();
            _worksheet = null;

            _workbook?.Close();
            _workbook = null;

            _excel?.Quit();
            _excel = null;

            GC.Collect();
            KillExcel();

            Closed?.Invoke();
        }
        private void KillExcel()
        {
            System.Diagnostics.Process[] PROC = System.Diagnostics.Process.GetProcessesByName("EXCEL");
            foreach (System.Diagnostics.Process PK in PROC)
            {
                if (PK.MainWindowTitle.Length == 0)
                {
                    PK.Kill();
                }
            }
        }
        /// <summary>
        /// Метод IDisposable для закрытия файла
        /// </summary>
        public void Dispose()
        {
            Close();
        }

        /// <summary>
        /// Получает значение клетки по ее расположению
        /// </summary>
        /// <param name="row">Строка (нумерация с 0)</param>
        /// <param name="column">Столбец (нумерация с 0)</param>
        /// <returns>Значение клетки</returns>
        public object GetCell(int row, int column)
        {
            return _worksheet?.Cells[++row, ++column].Value2 ?? "";
        }
        /// <summary>
        /// Присваивает выбранной клетке новое значение
        /// </summary>
        /// <param name="row">Строка (нумерация с 0)</param>
        /// <param name="column">Столбец (нумерация с 0)</param>
        /// <param name="data">Новое значение</param>
        public void SetCell(int row, int column, object data)
        {
            try
            {
                _worksheet.Cells[++row, ++column].Value2 = data;
            }
            catch (Exception e)
            {
                Logger.Log(e);
            }
        }
    }
}
