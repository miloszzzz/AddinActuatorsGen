using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpennessConsole.Models
{
    public class SubLevel
    {
        private int level = 0;

        public SubLevel() { }

        public string GetTabs()
        {
            string tabs = string.Empty;
            for (int i = 0; i < level; i++)
            {
                tabs += "\t";
            }
            return tabs;
        }

        public void NextLeveL() => level++;

        public void PrevLeveL() => level = level > 0 ? level - 1 : level;
        
        
    }

    
}
