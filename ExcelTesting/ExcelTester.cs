using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using ExcelHelper;
using System.Collections.Generic;

namespace ExcelTesting
{
    [TestClass]
    public class ExcelTester
    {
        [TestMethod]
        public void CalculateTest()
        {
            ExcelWorker ew = new ExcelWorker();
            PersonResult pr = ew.Calculate("bob", "beeb", new System.Collections.Generic.List<WordInfo>
            {
                ew.GetWord("цифра"      , "клетка"),
                ew.GetWord("цирк"       , "клетка"),
                ew.GetWord("балл"       , "порог"),
                ew.GetWord("аттестат"   , "порог"),
                ew.GetWord("бумага"     , "лист"),
                ew.GetWord("весна"      , "лист"),
                ew.GetWord("брелок"     , "ключ"),
                ew.GetWord("вода"       , "ключ"),
                ew.GetWord("англия"     , "язык"),
                ew.GetWord("вкус"       , "язык"),
                ew.GetWord("выстрел"    , "дробь"),
                ew.GetWord("делитель"   , "дробь"),
            });

            foreach (var i in pr.ToStringArray())
            {
                Console.Write(i + "\t");
            }
            ew.Close();
        }

        [TestMethod]
        public void InputTest()
        {
            ExcelWorker ew = new ExcelWorker();
            //ew.Close();
            //ew.InputPhase();

            List<WordInfo> ws = new List<WordInfo>();

            var wi = ew.GetWord("цифра", "клетка");
            ew.AddWord(wi);
            ws.Add(wi);
            wi = ew.GetWord("цирк", "клетка");
            ew.AddWord(wi);
            ws.Add(wi);
            wi = ew.GetWord("балл", "порог");
            ew.AddWord(wi);
            ws.Add(wi);
            wi = ew.GetWord("аттестат", "порог");
            ew.AddWord(wi);
            ws.Add(wi);
            wi = ew.GetWord("бумага", "лист");
            ew.AddWord(wi);
            ws.Add(wi);
            wi = ew.GetWord("весна", "лист");
            ew.AddWord(wi);
            ws.Add(wi);
            wi = ew.GetWord("брелок", "ключ");
            ew.AddWord(wi);
            ws.Add(wi);
            wi = ew.GetWord("вода", "ключ");
            ew.AddWord(wi);
            ws.Add(wi);
            wi = ew.GetWord("англия", "язык");
            ew.AddWord(wi);
            ws.Add(wi);
            wi = ew.GetWord("вкус", "язык");
            ew.AddWord(wi);
            ws.Add(wi);
            wi = ew.GetWord("выстрел", "дробь");
            ew.AddWord(wi);
            ws.Add(wi);
            wi = ew.GetWord("делитель", "дробь");
            ew.AddWord(wi);
            ws.Add(wi);

            ew.ResultReferencePhase();

            ew.SaveResultRef(new PersonResult { Name = "bob2", Group = "beebe2", FAss = 2, FSem = 2, Speed = 2, Originality = 1 }, ws);

            //foreach (var i in pr.ToStringArray())
            //{
            //    Console.Write(i + "\t");
            //}
            ew.Close();
        }

        [TestMethod]
        public void ResultTest()
        {
            ExcelWorker ew = new ExcelWorker();

            ew.ResultPhase();

            ew.SaveAllResults();

            ew.Close();
        }
    }
}
