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
}
