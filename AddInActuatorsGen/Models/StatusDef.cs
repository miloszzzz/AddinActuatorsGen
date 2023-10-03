using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AddInActuatorsGen.Models
{
    public class StatusDef
    {
        public int Status {  get; set; }
        public string Name { get; set; }

        public StatusDef(int status, string name)
        {
            Status = status;
            Name = name;
        }
    }
}
