using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace OpennessConsole.Enums
{
    /// <summary>
    /// Default Philogic groups
    /// </summary>
    public enum DefGroup
    {
        [Description("!!!!!Functions")]
        Functions,

        [Description("!!!!Handling")]
        Handling,

        [Description("!!!Devices")]
        Devices,

        [Description("!!Product")]
        Product,

        [Description("!Sequences")]
        Sequences
    }
}
