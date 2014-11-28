using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    public interface ICOMInterop_MainLine
    {
        Symbol Symbols { get; }
        Connectline Connectlines { get; }

    }
}
