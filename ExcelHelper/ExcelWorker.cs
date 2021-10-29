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
    public class ExcelWorker : ExcelWorkerBase
    {
        public ExcelWorker(string path) : base(path) { }
        public ExcelWorker() : base() { }
    }
}
