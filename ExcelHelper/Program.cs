using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelHelper
{
    //"C:\\Users\\Pepe\\Desktop\\Частоты1.xlsx"
    class Program
    {
        static string MakePath(string path)
        {
            return Directory.GetCurrentDirectory() + "\\" + path;
        }
        static void Main(string[] args)
        {
            //Console.WriteLine($"Current directory is '{Environment.CurrentDirectory}'");

            //string path = Path.GetFullPath(@"..\..\ExcelTables\Частоты.xlsx");
            //Console.WriteLine($"'..\\Debug' resolves to {path}");

            //Console.ReadKey();


            using (ExcelWorker ew2 = new ExcelWorker(Path.GetFullPath(@"..\..\ExcelTables\Частоты.xlsx")))
            {
                ew2.AddWord(new WordInfo
                {
                    Association = "клетка",
                    Word = "буба",
                    Frequency = 1,
                    FSem = 1,
                    FAss = 1
                });
            }
        }
    }
}
