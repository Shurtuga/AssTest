﻿using System.Collections.Generic;
using AssociationTestVisual.VisualTabs;

namespace AssociationTestVisual
{
    public static class GLOBALS
    {
        public static ExcelHelper.ExcelWorker Eww { get; set; }
        public static ExcelHelper.PersonResult GetPerson { get; set; }

        public static VisualTabs.WordsList Words { get; set; }
        public static List<ExcelHelper.WordInfo> WordInfos { get; set; }
        public static GroupsList Groups { get; set; }
    }
}
