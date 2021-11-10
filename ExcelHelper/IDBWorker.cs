using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    interface IDBWorker
    {
        /// <summary>
        /// Получает данные о слове из таблицы
        /// </summary>
        /// <param name="word">Слово</param>
        /// <returns>Информация о слове</returns>
        WordInfo GetWord(string word, string association);
        /// <summary>
        /// Добавляет новую ассоциацию в таблицу
        /// </summary>
        /// <param name="result">Параметры для слова</param>
        void AddWord(WordInfo info);
        /// <summary>
        /// Новая запись в таблицу результатов
        /// </summary>
        /// <param name="result">Параметры для человека</param>
        void SaveResult(PersonResult result);
        PersonResult Calculate(PersonResult[] results);
    }
}
