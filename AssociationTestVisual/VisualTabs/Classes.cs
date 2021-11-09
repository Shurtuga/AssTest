using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Xml.Serialization;

namespace AssociationTestVisual.VisualTabs
{
    public class GroupsList
    {
        public List<string> List;

        public GroupsList()
        {
            List = new List<string>();
        }
        public void Load() 
        {
            try
            {
                using (var stream = new FileStream("Groups.xml", FileMode.OpenOrCreate))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(List<string>));
                    List = (List<string>)serializer.Deserialize(stream);
                }
            }
            catch { }
        }

        public void Save()
        {
            using (var stream = new FileStream("Groups.xml", FileMode.OpenOrCreate))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(List<string>));
                serializer.Serialize(stream, List);
            }
        }
    }

    public class AssociationWord
    {
        public string Word;
        public List<string> Meanings = new List<string>();
    }

    public class WordsList
    {
        public static List<string> assTypes = new List<string>() {"Ассоциации по сходству", "Ассоциация по контрасту", 
            "Ассоциация по смежности в пространстве или времени", "Причинно-следственная ассоциация", "Ассоциация целое-часть",
            "Ассоциация-определение", "Сложнопонимаемая ассоциация" };
        public List<AssociationWord> semanticMeanings;
        public WordsList()
        {
            semanticMeanings= new List<AssociationWord>();
        }

        public void Save()
        {
            using (var stream = new FileStream("Words.xml", FileMode.OpenOrCreate))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(WordsList), new[] {typeof(AssociationWord) } );
                serializer.Serialize(stream, this);
            }
        }
        public void Load()
        {
            using (var stream = new FileStream("Words.xml", FileMode.OpenOrCreate))
            {
                XmlSerializer serializer = new XmlSerializer(typeof(WordsList), new[] { typeof(AssociationWord) });
                WordsList wl = (WordsList)serializer.Deserialize(stream);
                semanticMeanings = wl.semanticMeanings;
            }
        }
    }
}
