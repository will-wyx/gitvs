using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExpressMaster
{
    public class ProfileEntity
    {
        public string Name { get; set; }
        public char PinyinInitials { get; set; }
        public ValuesEntity[] Values { get; set; }
    }

    public class ValuesEntity
    {
        public Data4Cfg[] Items { get; set; }
    }
}
