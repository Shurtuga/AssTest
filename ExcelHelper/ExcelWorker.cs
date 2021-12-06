using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using System.IO;

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

        string freqBookPath = System.Configuration.ConfigurationManager.AppSettings["freqBookPath"];
        string resBookPath = System.Configuration.ConfigurationManager.AppSettings["resBookPath"];
        string wordsRefPath = System.Configuration.ConfigurationManager.AppSettings["wordsRefPath"];
        string resultsRefPath = System.Configuration.ConfigurationManager.AppSettings["resultsRefPath"];



        string _resultsBook;
        string _freqBook;

        ExcelWorkerBase _ewbFreq;
        ExcelWorkerBase _ewbRes;
        ExcelWorkerBase _ewbTmp;

        #region ctors
        public ExcelWorker(string path)
        {
            _resultsBook = System.IO.Path.GetFullPath(resBookPath);
            _freqBook = System.IO.Path.GetFullPath(freqBookPath);

            _ewbFreq = new ExcelWorkerBase(_freqBook);
            _ewbRes = new ExcelWorkerBase(_resultsBook);
        }
        public ExcelWorker()
        {
            _resultsBook = System.IO.Path.GetFullPath(resBookPath);
            _freqBook = System.IO.Path.GetFullPath(freqBookPath);

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

        #region phases
        public void InputPhase()
        {
            _ewbTmp.Close();
            _ewbRes.Open(Path.GetFullPath(wordsRefPath));
            _ewbFreq.Open(Path.GetFullPath(freqBookPath));
        }
        public void ResultReferencePhase()
        {
            _ewbRes.Open(Path.GetFullPath(resultsRefPath));
            _ewbFreq.Open(Path.GetFullPath(freqBookPath));
        }
        public void ResultPhase()
        {
            //await Task.Factory.StartNew(() => {
            //    _ewbTmp = new ExcelWorkerBase(resBookPath);
            //});
            _ewbTmp = new ExcelWorkerBase(resBookPath);
            _ewbRes.Open(Path.GetFullPath(resultsRefPath));
            _ewbFreq.Open(Path.GetFullPath(wordsRefPath));
        }
        #endregion

        #region DBWorker Methods
        public WordInfo GetWord(string word, string association)
        {
            return getWord(_ewbRes, word, association);
        }
        private WordInfo getWord(ExcelWorkerBase ewb, string word, string association)
        {
            //Open(_freqBook);
            ewb.SelectSheet(association);

            int c = 0;
            string w;
            //while ((w = GetCell(c, 0).ToString().Trim(' '))!= word || (w == ""))
            //{
            //    c++;
            //}

            for (; ; )
            {
                w = ewb.GetCell(c, 0).ToString().Trim(' ');
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
            addWord(_ewbRes, info);
            addWord(_ewbFreq, info);
        }
        private void addWord(ExcelWorkerBase ewb, WordInfo info)
        {
            //Open(_freqBook);
            ewb.SelectSheet(info.Association);

            int c = 0;
            string w;

            for (; ; )
            {
                w = ewb.GetCell(c, 0).ToString().Trim(' ');
                if (w == info.Word)
                {
                    int t = int.Parse(ewb.GetCell(c, 1).ToString());
                    ewb.SetCell(c, 1, t + 1);
                    break;
                }
                else if (w == "")
                {
                    ewb.AddRow(0, info.ToArray());
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
        public void SaveResultRef(PersonResult result, List<WordInfo> words)
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

            string[] ws =
            {
                "Клетка",
                "Лист",
                "Дробь",
                "Ключ",
                "Порог",
                "Язык",
            };


            string[] res = new string[10];

            res[0] = result.Name;
            res[1] = result.Speed.ToString();
            res[2] = result.FAss.ToString();
            res[3] = result.FSem.ToString();

            for (int i = 0; i < 6; i++)
            {
                System.Text.StringBuilder sb = new StringBuilder();

                for (int j = 0; j < words.Count; j++)
                {
                    string t;
                    if ((t = words[j].Association) == ws[i])
                    {
                        sb.Append(t + "@");
                    }
                }

                res[4 + i] = sb.ToString();
            }

            _ewbRes.AddRow(0, res);
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
