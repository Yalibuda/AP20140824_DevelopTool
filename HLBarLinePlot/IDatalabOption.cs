using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.HLBarLinePlot
{
    public interface IDatalabOption
    {
        bool AutoDecimal { get; set; }
        int DecimalPlace { get; set; }
    }
}
