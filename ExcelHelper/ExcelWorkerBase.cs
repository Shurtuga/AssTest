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
    public class ExcelWorkerBase : IDisposable
    {
        internal Application _excel;
        internal Workbook _workbook;
        internal Worksheet _worksheet;
        internal string _path;
        internal string _password = "59Kk{Mu.@j";

        public event System.Action Closing;
        public event System.Action Closed;
        public event System.Action Opened;
        public event System.Action Loaded;
        public event System.Action SheetChanged;

        #region Ctors
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
                this._path = path;

                Loaded?.Invoke();
            }
            catch (Exception e)
            {
                //Logger.Log(e);
                NewFile(path);
                //Close();
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
        #endregion

        #region protect/unprotect
        /// <summary>
        /// Защищает вкладку от изменений
        /// </summary>
        public void ProtectSheet()
        {
            _worksheet.Protect();
        }
        /// <summary>
        /// Защищает все вкладки в документе
        /// </summary>
        public void ProtectAllSheets()
        {
            foreach (Worksheet sheet in _workbook.Worksheets)
            {
                sheet.Protect(_password);
            }
        }
        /// <summary>
        /// Защищает вкладку от изменений
        /// </summary>
        /// <param name="password">Пароль</param>
        public void ProtectSheet(string password)
        {
            _worksheet.Protect(password);
        }
        /// <summary>
        /// Разрешает изменения во вкладке
        /// </summary>
        public void UnprotectSheet()
        {
            _worksheet.Unprotect();
        }
        /// <summary>
        /// Разрешает изменения во всех вкладках
        /// </summary>
        public void UnprotectAllSheets()
        {
            foreach (Worksheet sheet in _workbook.Worksheets)
            {
                sheet.Unprotect(_password);
            }
        }
        /// <summary>
        /// Разрешает изменения во вкладке
        /// </summary>
        /// <param name="password">Пароль</param>
        public void UnprotectSheet(string password)
        {
            _worksheet.Unprotect(password);
        }
        #endregion

        #region save(as)
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
            this._path = path;
        }
        #endregion

        #region create files
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
        #endregion

        #region get create select and delete sheets
        /// <summary>
        /// Создает новую вкладку в документе на первом месте
        /// </summary>
        public void NewSheet()
        {
            //Worksheet tmp = _workbook.Worksheets.Add(After: _workbook.Worksheets[_workbook.Worksheets.Count]);
            Worksheet _ = _workbook.Worksheets.Add();
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
            try
            {
                _worksheet = _workbook.Worksheets[name];
                _worksheet.Activate();
            }
            catch (Exception e)
            {
                NewSheet(name);
                throw;
            }

            SheetChanged?.Invoke();
        }
        /// <summary>
        /// Открывает вкладку определенного номера
        /// </summary>
        /// <param name="sheet">Номер вкладки (нумерация с 0)</param>
        public void SelectSheet(int sheet)
        {
            try
            {
                _worksheet = _workbook.Sheets[++sheet];
                _worksheet.Activate();
            }
            catch (Exception e)
            {
                NewSheet();
            }

            SheetChanged?.Invoke();
        }

        /// <summary>
        /// Возвращает список всех названий вкладок в файле
        /// </summary>
        /// <returns>Список всех названий вкладок в файле</returns>
        public string[] GetSheetNames()
        {
            List<string> sh = new List<string>();
            foreach (Worksheet ws in _workbook.Worksheets)
            {
                sh.Add(ws.Name);
            }

            return sh.ToArray();
        }
        #endregion

        #region open close dispose
        /// <summary>
        /// Открывает файл по его пути
        /// </summary>
        /// <param name="path">Путь файла</param>
        public void Open(string path)
        {
            try
            {
                _workbook = _excel.Workbooks.Open(path);
                SelectSheet(0);
            }
            catch (Exception)
            {
                NewFile(path);
            }

            this._path = path;

            Opened?.Invoke();
        }
        /// <summary>
        /// Закрывает файл с сохранением изменений
        /// </summary>
        public void Close()
        {
            Closing?.Invoke();

            ((Excel._Workbook)_workbook)?.Save();
            _worksheet = null;

            _workbook?.Close();
            _workbook = null;

            _excel?.Quit();
            _excel = null;

            //GC.Collect();
            //KillExcel();

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
            Console.WriteLine("DISPOSED");
            Close();
        }
        #endregion

        #region get/set cell
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
                //Logger.Log(e);
            }
        }
        #endregion

        #region get/set range
        /// <summary>
        /// Читает массив клеток в заданном диапазоне
        /// </summary>
        /// <param name="startx">Строка начала</param>
        /// <param name="starty">Строка окончания</param>
        /// <param name="endx">Столбец начала</param>
        /// <param name="endy">Столбец окончания</param>
        /// <returns>Двумерный массив значений клеток</returns>
        public object[,] GetRange(int startx, int starty, int endx, int endy)
        {
            Range rg = (Range)_worksheet.Range[_worksheet.Cells[++startx, ++starty], _worksheet.Cells[++endx, ++endy]];

            return rg.Value2;
        }
        public string[,] GetRangeString(int startx, int starty, int endx, int endy)
        {
            object[,] tmp = GetRange(startx, starty, endx, endy);

            string[,] res = new string[endx - startx + 1, endy - starty + 1];
            for (int i = 0; i <= endx - startx; i++)
            {
                for (int j = 0; j <= endy - starty; j++)
                {
                    res[i, j] = tmp[i + 1, j + 1]?.ToString() ?? "";
                }
            }
            return res;
        }
        /// <summary>
        /// Присваивает новые значения клеткам в заданном диапазоне
        /// </summary>
        /// <param name="startx">Строка начала</param>
        /// <param name="starty">Строка окончания</param>
        /// <param name="endx">Столбец начала</param>
        /// <param name="endy">Столбец окончания</param>
        /// <param name="data">Матрица значений</param>
        public void SetRange(int startx, int starty, int endx, int endy, object[,] data)
        {
            Range rg = (Range)_worksheet.Range[_worksheet.Cells[++startx, ++starty], _worksheet.Cells[++endx, ++endy]];

            rg.Value2 = data;
        }
        /// <summary>
        /// Присваивает новые значения клеткам в заданном диапазоне
        /// </summary>
        /// <param name="startx">Строка начала</param>
        /// <param name="starty">Столбец окончания</param>
        /// <param name="data">Матрица значений</param>
        public void SetRange(int startx, int starty, object[,] data)
        {
            int endx = data.GetLength(0);
            int endy = data.GetLength(1);
            Range rg = (Range)_worksheet.Range[_worksheet.Cells[startx + 1, starty + 1], _worksheet.Cells[endx + startx, endy + starty]];

            rg.Value2 = data;
        }
        #endregion

        #region get/set row
        /// <summary>
        /// Добавляет новую строку в конец заданного столбца
        /// </summary>
        /// <param name="column">Столбец</param>
        /// <param name="data">Строка</param>
        public void AddRow(int column, object[] data)
        {
            int c = 1;
            //var h = ;
            while (_worksheet.Cells[c, column + 1].Value2 != null)
            {
                c++;
            }
            int x = data.GetLength(0);

            Range rg = (Range)_worksheet.Range[_worksheet.Cells[c, column + 1], _worksheet.Cells[c, column + x]];

            rg.Value2 = data;
        }
        /// <summary>
        /// Вставляет строку
        /// </summary>
        /// <param name="column">Строка начала</param>
        /// <param name="row">Столбец начала</param>
        /// <param name="data">Строка для вставки</param>
        public void InsertRow(int column, int row, object[] data)
        {
            InsertBlankRow(row);

            int x = data.GetLength(0);
            Range rg = (Range)_worksheet.Range[_worksheet.Cells[row + 1, column + 1], _worksheet.Cells[row + 1, column + x]];

            rg.Value2 = data;
        }
        private void InsertBlankRow(int row)
        {
            Range _range = (Range)_worksheet.Rows[row + 1, Type.Missing];
            _range.Select();
            _range.Insert(XlInsertShiftDirection.xlShiftDown, Type.Missing);
        } 
        #endregion
    }
}
