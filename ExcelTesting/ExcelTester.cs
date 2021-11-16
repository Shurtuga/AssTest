using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using ExcelHelper;

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
    }
}
