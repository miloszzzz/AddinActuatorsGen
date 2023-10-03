using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddInActuatorsGen.Models
{
    public class TextlistEntry
    {
        public int Number { get; set; }
        public List<string> Texts { get; set; }
        
        public TextlistEntry(int culturesNumber)
        {
            Texts = new List<string>(culturesNumber);
        }
    }
}
