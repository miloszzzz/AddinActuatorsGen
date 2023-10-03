using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Siemens.Engineering.Hmi.Tag;
using Siemens.Engineering.SW.Tags;


namespace AddInActuatorsGen.Models
{
    public class Actuator
    {
        public string Name { get; set; }
        public int Number { get; set; }
        public EnumStations Station { get; set; }
        public int Constant { get; set; }
        public string InputRetract { get; set; }
        public string InputExtend { get; set; }
        public string OutputRetract { get; set; }
        public string OutputExtend { get; set;}

        public Actuator() 
        {
            Name = "Unknown";
            Number = 0;
            Station = EnumStations.st1;
            Constant = 0;
            InputRetract = "Unknown";
            InputExtend = "Unknown";
            OutputRetract = "Unknown";
            OutputExtend = "Unknown";
        }

        public Actuator(string name, int number, EnumStations station, int constant, string inputRetract, string inputExtend, string outputRetract, string outputExtend)
        {
            Name = name;
            Number = number;
            Station = station;
            Constant = constant;
            InputRetract = inputRetract;
            InputExtend = inputExtend;
            OutputRetract = outputRetract;
            OutputExtend = outputExtend;
        }
    }
}
