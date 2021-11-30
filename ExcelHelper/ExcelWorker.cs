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
    public class ExcelWorker : IDBWorker, IDisposable
    {
        //string _resultsBook = System.IO.Path.GetFullPath(@"..\..\ExcelTables\Результаты.xlsx");
        //string _freqBook = System.IO.Path.GetFullPath(@"..\..\ExcelTables\Частоты.xlsx");

        //string _resultsBook = System.IO.Path.GetFullPath(
        //    @"C:\Users\User\source\repos\associationstest\ExcelHelper\ExcelTables\Результаты.xlsx");
        //string _freqBook = System.IO.Path.GetFullPath(
        //    @"C:\Users\User\source\repos\associationstest\ExcelHelper\ExcelTables\Частоты.xlsx");

        string _resultsBook;
        string _freqBook;

        ExcelWorkerBase _ewbFreq;
        ExcelWorkerBase _ewbRes;

        #region ctors
        public ExcelWorker(string path)
        {
            _resultsBook = System.Configuration.ConfigurationManager.AppSettings["resBookPath"];
            _freqBook = System.Configuration.ConfigurationManager.AppSettings["freqBookPath"];

            _ewbFreq = new ExcelWorkerBase(_freqBook);
            _ewbRes = new ExcelWorkerBase(_resultsBook);
        }
        public ExcelWorker()
        {
            _resultsBook = System.Configuration.ConfigurationManager.AppSettings["resBookPath"];
            _freqBook = System.Configuration.ConfigurationManager.AppSettings["freqBookPath"];

            _ewbFreq = new ExcelWorkerBase(System.IO.Path.GetFullPath(_freqBook));
            _ewbRes = new ExcelWorkerBase(System.IO.Path.GetFullPath(_resultsBook));
        }
        #endregion

        #region !!!need to be removed
        public void Close()
        {
            _ewbFreq.Close();
            _ewbRes.Close();

            GC.Collect();
            KillExcel();
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
        public void Dispose()
        {
            Close();
        }
        #endregion

        #region DBWorker Methods
        public WordInfo GetWord(string word, string association)
        {
            //Open(_freqBook);
            _ewbFreq.SelectSheet(association);

            int c = 0;
            string w;
            //while ((w = GetCell(c, 0).ToString().Trim(' '))!= word || (w == ""))
            //{
            //    c++;
            //}

            for (; ; )
            {
                w = _ewbFreq.GetCell(c, 0).ToString().Trim(' ');
                if (w == word)
                {
                    break;
                }
                else if (w == "")
                {
                    return new WordInfo
                    {
                        Association = association,
                        Word = word,
                        Frequency = 1,
                        FAss = -1,
                        FSem = -1
                    };
                }
                c++;
            }
            return ParseRow(c);
        }

        WordInfo ParseRow(int row)
        {
            return new WordInfo()
            {
                Association = _ewbFreq._worksheet.Name,
                Word = _ewbFreq.GetCell(row, 0).ToString(),
                Frequency = int.Parse(_ewbFreq.GetCell(row, 1).ToString()),
                FSem = int.Parse(_ewbFreq.GetCell(row, 2).ToString()),
                FAss = int.Parse(_ewbFreq.GetCell(row, 3).ToString())
            };
        }

        public void AddWord(WordInfo info)
        {
            //Open(_freqBook);
            _ewbFreq.SelectSheet(info.Association);

            int c = 0;
            string w;

            for (; ; )
            {
                w = _ewbFreq.GetCell(c, 0).ToString().Trim(' ');
                if (w == info.Word)
                {
                    int t = int.Parse(_ewbFreq.GetCell(c, 1).ToString());
                    _ewbFreq.SetCell(c, 1, t + 1);
                    break;
                }
                else if (w == "")
                {
                    _ewbFreq.AddRow(0, info.ToArray());
                    break;
                }
                c++;
            }
        }

        public void SaveResult(PersonResult result)
        {
            //Open(_resultsBook);
            try
            {
                _ewbRes.SelectSheet(result.Group);
            }
            catch (Exception)
            {
                _ewbRes.NewSheet(result.Group);
                _ewbRes.SelectSheet(result.Group);
                _ewbRes.AddRow(0, new object[] { "Имя", "Беглость", "Оригинальность", "ГСем", "ГАсс" });
            }

            _ewbRes.AddRow(0, result.ToStringArray());
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

        public PersonResult Calculate(string name, string group, List<WordInfo> results)
        {
            Dictionary<string, List<int>[]> raspr = new Dictionary<string, List<int>[]>();

            foreach (var r in results)
            {
                string t = r.Association;

                if (raspr.ContainsKey(t))
                {
                    int a1 = r.FAss, a2 = r.FSem, a3 = r.Frequency;

                    if (!raspr[t][0].Contains(a1))
                    {
                        raspr[t][0].Add(a1);
                    }
                    if (!raspr[t][1].Contains(a2))
                    {
                        raspr[t][1].Add(a2);
                    }
                    raspr[t][2].Add(a3);
                    raspr[t][3].Add(1);
                }
                else
                {
                    raspr.Add(t, new List<int>[]
                    {
                        new List<int>{r.FAss},
                        new List<int>{r.FSem},
                        new List<int>{r.Frequency},
                        new List<int>{1}
                    });
                }
            }

            int cnt = 0;
            foreach (var i in raspr)
            {
                cnt += (i.Value[0]).Count;
            }

            double fass = (double)cnt / raspr.Count;

            cnt = 0;
            foreach (var i in raspr)
            {
                cnt += (i.Value[1]).Count;
            }

            double fsem = (double)cnt / raspr.Count;

            cnt = 0;
            foreach (var i in raspr)
            {
                cnt += (i.Value[3]).Count;
            }

            double speed = (double)cnt / raspr.Count;

            double orig = 0;
            foreach (var i in raspr)
            {
                double origt = 0;
                foreach (var j in i.Value[2])
                {
                    origt += j;
                }
                orig += origt / i.Value[2].Count;
            }

            orig /= raspr.Count;
            orig = Math.Round(orig);

            return new PersonResult
            {
                Name = name,
                Group = group,
                Speed = speed,
                FAss = fass,
                FSem = fsem,
                Originality = orig
            };
        }
        #endregion

        #region Async Methods

        public async Task<WordInfo> GetWordAsync(string word, string association)
        {
            WordInfo res = null;

            await Task.Factory.StartNew(() =>
            {
                res = GetWord(word, association);
            });

            return res;
        }

        public async Task AddWordAsync(WordInfo info)
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

        //~ExcelWorker()
        //{
        //    Console.WriteLine("ENDING HERE");
        //    _ewbFreq.Close();
        //    _ewbRes.Close();
        //}
    }
}
