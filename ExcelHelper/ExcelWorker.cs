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
        string _freqBook = System.IO.Path.GetFullPath(@"..\..\ExcelTables\Частоты.xlsx");

        public ExcelWorker(string path) : base(path) { }
        public ExcelWorker() : base() { }

        #region DBWorker Methods
        public WordInfo GetWord(string word, string association)
        {
            Open(_freqBook);
            SelectSheet(association);

            int c = 0;
            string w;
            //while ((w = GetCell(c, 0).ToString().Trim(' '))!= word || (w == ""))
            //{
            //    c++;
            //}

            for(; ; )
            {
                w = GetCell(c, 0).ToString().Trim(' ');
                if (w == word)
                {
                    break;
                }
                else if (w == "")
                {
                    return null;
                }
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

        public void AddWord(WordInfo info)
        {
            Open(_freqBook);
            SelectSheet(info.Association);

            int c = 0;
            string w;

            for (; ; )
            {
                w = GetCell(c, 0).ToString().Trim(' ');
                if (w == info.Word)
                {
                    int t = int.Parse(GetCell(c, 1).ToString());
                    SetCell(c, 1, t + 1);
                    break;
                }
                else if (w == "")
                {
                    AddRow(0, info.ToArray());
                    break;
                }
            }
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

        public PersonResult Calculate(PersonResult[] results)
        {
            double fass = 0, fsem = 0, orig = 0;

            foreach (var r in results)
            {
                orig += r.Originality;
                fass += r.FAss;
                fsem += r.FSem;
            }

            orig /= 6;
            fass /= 6;
            fsem /= 6;

            return new PersonResult
            {
                Name = results[0].Name,
                Group = results[0].Group,
                Originality = orig,
                FAss = fass,
                FSem = fsem
            };
        } 
        #endregion

        #region Async Methods

        public async Task<WordInfo> GetInfoAsync(string word, string association)
        {
            WordInfo res = null;

            await Task.Factory.StartNew(() =>
            {
                res = GetWord(word, association);
            });

            return res;
        }

        public async Task AddWordEntryAsync(WordInfo info)
        {
            await Task.Factory.StartNew(() =>
            {
                AddWord(info);
            });
        }

        public async Task SaveResultAsync(PersonResult result)
        {
            await Task.Factory.StartNew(() =>
            {
                SaveResult(result);
            });
        }

        #endregion
    }
}
