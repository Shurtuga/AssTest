//using System;
//using System.Collections.Generic;
//using System.Linq;
//using System.Text;
//using System.Threading.Tasks;
//using Microsoft.Office.Interop.Excel;
//using Excel = Microsoft.Office.Interop.Excel;

//namespace ExcelHelper
//{
//    class Program
//    {
//        static void Main(string[] args)
//        {
//            Application excel = new Application();
//            Workbook wb = null;
//        }
//    }
//}

using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace ExcelHelper
{
    class Program
    {
        static string MakePath(string path)
        {
            return Directory.GetCurrentDirectory() + "\\" + path;
        }
        static void Main(string[] args)
        {
            //ExcelWorker ew = new ExcelWorker();
            using (ExcelWorker ew = new ExcelWorker())
            {
                ew.Open(MakePath("test1.xlsx"));
                ew.NewSheet("aboba2");
                //ew.SelectSheet("Лист7");

                //ew.SaveAs(); 
            }
            //ew.Close();

            //ExcelWorker ew = new ExcelWorker(Directory.GetCurrentDirectory() + "\\" + "test.xlsx");

            //ew.Open(0);

            //string[,] res = new string[5, 5];
            //for (int i = 0; i < 5; i++)
            //{
            //    for (int j = 0; j < 5; j++)
            //    {
            //        res[i, j] = ew.GetCell(i, j).ToString();
            //        Console.Write(res[i,j] + "\t");
            //    }
            //    Console.WriteLine();
            //}


            //Console.WriteLine(ew.GetCell(0, 0).ToString());
            //ew.SetCell(1, 0, "test01");

            //ew.Close();
            

            
            //Console.ReadKey();
        }
        private static void KillExcel()
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
    }
}
