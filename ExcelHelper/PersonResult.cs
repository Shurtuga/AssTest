namespace ExcelHelper
{
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
}
