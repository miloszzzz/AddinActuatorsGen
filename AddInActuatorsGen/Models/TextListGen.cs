using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddInActuatorsGen.Models
{
    public class TextListGen
    {
        public string Name { get; set; }
        public List<string> Comment { get; set; }
        public Dictionary<int, TextlistEntry> Entries { get; set; }
        public int CulturesNumber { get; set; }

        public TextListGen(string name, int culturesNumber) 
        { 
            Name = name;
            Comment = new List<string>(culturesNumber);
            Entries = new Dictionary<int, TextlistEntry>();
            CulturesNumber = culturesNumber;
        }
    }
}
