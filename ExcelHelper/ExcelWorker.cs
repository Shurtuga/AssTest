using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

using System.IO;
using System.Threading;

namespace ExcelHelper
{
    /// <summary>
    /// Класс для работы с Microsoft Excel
    /// </summary>
    public class ExcelWorker : IDBWorker, IDisposable
    {
        public Dictionary<string, List<string>> Words { get; private set; }
        public Dictionary<string, List<string>> ResRefs { get; private set; }

        public event EventHandler ExcelLoaded;

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


        ExcelWorkerBase _ewb;
        //ExcelWorkerBase _ewbRes;

        #region ctors
        public ExcelWorker()
        {
            try
            {
                var jsonString = File.ReadAllText(Path.GetFullPath(freqBookPath));
                Words = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(jsonString);
            }
            catch (Exception)
            {
                Words = new Dictionary<string, List<string>>();
                string[] ws =
                {
                "клетка",
                "лист",
                "дробь",
                "ключ",
                "порог",
                "язык",
                };

                foreach (var w in ws)
                {
                    Words.Add(w, new List<string>(64));
                }
            }

            try
            {
                var jsonString = File.ReadAllText(Path.GetFullPath(resultsRefPath));
                ResRefs = JsonConvert.DeserializeObject<Dictionary<string, List<string>>>(jsonString);
            }
            catch (Exception)
            {
                ResRefs = new Dictionary<string, List<string>>();
            }
        }
        #endregion

        #region !!!closing
        public void SaveJson()
        {
            var jsonString = JsonConvert.SerializeObject(Words, Formatting.Indented);
            File.WriteAllText(Path.GetFullPath(freqBookPath), jsonString);

            jsonString = JsonConvert.SerializeObject(ResRefs, Formatting.Indented);
            File.WriteAllText(Path.GetFullPath(resultsRefPath), jsonString);
        }
        public void Close()
        {
            SaveJson();

            _ewb.Close();
            //_ewbRes.Close();
            //_ewbTmp?.Close();

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
            if (_ewb == null)
            {
                _ewb = new ExcelWorkerBase(Path.GetFullPath(wordsRefPath)); 
            }
            else
            {
                _ewb.Open(wordsRefPath);
            }
            ExcelLoaded?.Invoke(this, EventArgs.Empty);
        }

        public async void InputPhaseAsync()
        {
            await Task.Factory.StartNew(InputPhase);
        }
        public void ResultReferencePhase()
        {
            
        }
        public void ResultPhase()
        {
            SaveJson();

            _ewb?.Save();

            _ewb?.Open(Path.GetFullPath(resBookPath));
        }
        #endregion

        #region DBWorker Methods
        public WordInfo GetWord(string word, string association)
        {
            return getWord(_ewb, word, association);
        }
        private WordInfo getWord(ExcelWorkerBase ewb, string word, string association)
        {
            try
            {
                ewb.SelectSheet(association);
            }
            catch (Exception)
            {
                ewb.SelectSheet(association);
                ewb.AddRow(0, new object[] { "Слово", "Частота", "ГСем", "ГАсс" });
            }

            int c = 0;
            string w;

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
            return ParseRow(ewb, association, c);
        }

        private WordInfo getWord(List<string> words, string word)
        {
            for (int i = 0; i < words.Count; i++)
            {
                string[] ws = words[i].Split(new char[] { '#' }, StringSplitOptions.RemoveEmptyEntries);

                if (ws[0] == word)
                {
                    return new WordInfo()
                    {
                        Word = ws[0],
                        Frequency = int.Parse(ws[1]),
                        FSem = int.Parse(ws[2]),
                        FAss = int.Parse(ws[3])
                    };
                }
            }
            return new WordInfo()
            {
                Word = word,
                Frequency = 1,
                FSem = -1,
                FAss = -1
            };
        }
        WordInfo ParseRow(ExcelWorkerBase ewb, string sheet, int row)
        {
            ewb.SelectSheet(sheet);
            return new WordInfo()
            {
                Association = ewb._worksheet.Name,
                Word = ewb.GetCell(row, 0).ToString().Trim(' '),
                Frequency = int.Parse(ewb.GetCell(row, 1).ToString()),
                FSem = int.Parse(ewb.GetCell(row, 2).ToString()),
                FAss = int.Parse(ewb.GetCell(row, 3).ToString())
            };
        }

        public void AddWord(WordInfo info)
        {
            var words = Words[info.Association.ToLower()];

            for (int i = 0; i < words.Count; i++)
            {
                string[] ws = words[i].Split(new char[] { '#' }, StringSplitOptions.RemoveEmptyEntries);

                if (ws[0] == info.Word)
                {
                    int freq = int.Parse(ws[1]);

                    ws[1] = (freq + 1).ToString();

                    Words[info.Association.ToLower()][i] = String.Join("#", ws);


                    return;
                }
            }

            info.Frequency = 1;
            Words[info.Association.ToLower()].Add(info.ToString());

            //addWord(_ewb, info);
        }
        private void addWord(ExcelWorkerBase ewb, WordInfo info)
        {
            try
            {
                ewb.SelectSheet(info.Association);
            }
            catch (Exception)
            {
                ewb.SelectSheet(info.Association);
                ewb.AddRow(0, new object[] { "Слово", "Частота", "ГСем", "ГАсс" });
            }

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
                    info.Frequency = 1;
                    ewb.AddRow(0, info.ToArray());
                    break;
                }
                c++;
            }
        }

        public void SaveResult(PersonResult result)
        {
            //try
            //{
            //    _ewbRes.SelectSheet(result.Group);
            //}
            //catch (Exception)
            //{
            //    _ewbRes.SelectSheet(result.Group);
            //    _ewbRes.AddRow(0, new object[] { "Имя", "Беглость", "Оригинальность", "ГСем", "ГАсс" });
            //}

            //_ewbRes.AddRow(0, result.ToStringArray());
        }
        public void SaveResultRef(PersonResult result, List<WordInfo> words)
        {
            if (!ResRefs.ContainsKey(result.Group))
            {
                ResRefs.Add(result.Group, new List<string>());
            }

            var pr = ResRefs[result.Group];

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

            StringBuilder sb = new StringBuilder();
            for (int i = 0; i < 6; i++)
            {
                for (int j = 0; j < words.Count; j++)
                {
                    if (words[j].Association.ToLower() == ws[i].ToLower())
                    {
                        sb.Append(words[j].Word.Trim(' ') + "@");
                    }
                }

                res[4 + i] = sb.ToString();
                sb.Clear();
            }

            sb.Clear();
            for (int i = 0; i < 10; i++)
            {
                sb.Append(res[i].Trim(' ') + "#");
            }

            pr.Add(sb.ToString());
        }

        public void SaveAllResults()
        {
            if (_ewb == null) return;

            foreach (var key in ResRefs.Keys)
            {
                try
                {
                    _ewb.SelectSheet(key);
                }
                catch (Exception)
                {
                    _ewb.SelectSheet(key);
                    _ewb.AddRow(0, new object[]
                    {
                        "Имя",
                        "Оригинальность",
                        "Беглость",
                        "ГСем",
                        "ГАсс"
                    });
                }

                var results = ResRefs[key];

                PersonResult res = new PersonResult();

                foreach (var r in ResRefs[key])
                {
                    string[] tmp = r.Split(new char[] { '#' }, StringSplitOptions.RemoveEmptyEntries);

                    res.Name = tmp[0];
                    res.Speed = int.Parse(tmp[1]);
                    res.FSem = int.Parse(tmp[2]);
                    res.FAss = int.Parse(tmp[3]);
                    res.Group = key;

                    List<WordInfo> words = new List<WordInfo>();

                    for (int i = 0; i < 6; i++)
                    {
                        string[] w = tmp[4 + i].Split(new char[] { '@' }, StringSplitOptions.RemoveEmptyEntries);
                        foreach (var v in w)
                        {
                            var wi = getWord(Words.ElementAt(i).Value, v);
                            wi.Association = Words.ElementAt(i).Key;
                            words.Add(wi);
                        }
                    }

                    PersonResult tmpres = Calculate(tmp[0], key, words);
                    _ewb.AddRow(0, tmpres.ToStringArray());
                }
            }
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

    }
}
