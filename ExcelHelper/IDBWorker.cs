using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelHelper
{
    public class WordInfo
    {
        public string Association { get; set; }
        public string Word { get; set; }
        public int Frequency { get; set; }
        public int FSem { get; set; }
        public int FAss { get; set; }

        public object[] ToArray()
        {
            return new object[] {Word, Frequency, FSem, FAss };
        }
    }
    /// <summary>
    /// Результаты тестирования для человека
    /// </summary>
    public class PersonResult
    {
        /// <summary>
        /// Группа
        /// </summary>
        public string Group { get; set; }
        /// <summary>
        /// Имя / никнейм
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// Беглость
        /// </summary>
        public int Speed { get; set; }
        /// <summary>
        /// Оригинальность
        /// </summary>
        public double Originality { get; set; }
        /// <summary>
        /// Гибкость семантическая
        /// </summary>
        public double FSem { get; set; }
        /// <summary>
        /// Гибкость ассоциативная
        /// </summary>
        public double FAss { get; set; }

        public string[] ToStringArray()
        {
            return new string[] { Name, Speed.ToString(), Originality.ToString(), FAss.ToString(), FSem.ToString() };
        }

    }
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
