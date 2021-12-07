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
            ws.Add(wi);
            wi = ew.GetWord("цирк", "клетка");
            ws.Add(wi);
            wi = ew.GetWord("балл", "порог");
            ws.Add(wi);
            wi = ew.GetWord("аттестат", "порог");
            ws.Add(wi);
            wi = ew.GetWord("бумага", "лист");
            ws.Add(wi);
            wi = ew.GetWord("весна", "лист");
            ws.Add(wi);
            wi = ew.GetWord("брелок", "ключ");
            ws.Add(wi);
            wi = ew.GetWord("вода", "ключ");
            ws.Add(wi);
            wi = ew.GetWord("англия", "язык");
            ws.Add(wi);
            wi = ew.GetWord("вкус", "язык");
            ws.Add(wi);
            wi = ew.GetWord("выстрел", "дробь");
            ws.Add(wi);
            wi = ew.GetWord("делитель", "дробь");
            ws.Add(wi);

            ew.SaveResultRef(new PersonResult { Name = "bob", Group = "beeb", FAss = 2, FSem = 2, Speed = 2, Originality = 1 }, ws);

            //foreach (var i in pr.ToStringArray())
            //{
            //    Console.Write(i + "\t");
            //}
            ew.Close();
        }
    }
}
