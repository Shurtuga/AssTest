using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    /// <summary>
    /// Класс для работы с Microsoft Excel
    /// </summary>
    public class ExcelWorker : ExcelWorkerBase, IDBWorker
    {
        string _resultsBook = System.IO.Path.GetFullPath(@"..\..\ExcelTables\Результаты.xlsx");
        public string _freqBook = System.IO.Path.GetFullPath(@"..\..\ExcelTables\Частоты.xlsx");



        public ExcelWorker(string path) : base(path) { }
        public ExcelWorker() : base() { }

        public WordInfo GetInfo(string word, string association)
        {
            Open(_freqBook);
            SelectSheet(association);

            int c = 0;
            while (GetCell(c, 0).ToString().Trim(' ') != word)
            {
                c++;
            }
            return ParseRow(c);
        }


        WordInfo ParseRow(int row)
        {
            return new WordInfo()
            {
                Association = _worksheet.Name,
                Word = GetCell(row, 0).ToString(),
                Frequency = int.Parse(GetCell(row, 1).ToString()),
                FSem = int.Parse(GetCell(row, 2).ToString()),
                FAss = int.Parse(GetCell(row, 3).ToString())
            };
        }

        public void AddWordEntry(WordInfo info)
        {
            Open(_freqBook);
            SelectSheet(info.Association);

            AddRow(0, info.ToArray());

            //int c = 0;
            //while (GetCell(c, 0).ToString().Trim(' ') != "")
            //{
            //    c++;
            //}

            //SetCell(c, 0, info.Word);
            //SetCell(c, 1, info.Frequency);
            //SetCell(c, 2, info.FSem);
            //SetCell(c, 3, info.FAss);
        }

        public void SaveResult(PersonResult result)
        {
            Open(_resultsBook);
            try
            {
                SelectSheet(result.Group);
            }
            catch (Exception)
            {
                NewSheet(result.Group);
                SelectSheet(result.Group);
                AddRow(0, new object[] { "Имя", "Беглость", "Оригинальность", "ГСем", "ГАсс" });
            }

            AddRow(0, result.ToStringArray());
        }
    }
}
