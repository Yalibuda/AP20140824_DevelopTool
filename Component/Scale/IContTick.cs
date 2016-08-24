using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Scale
{
    public interface IContTick
    {
        int NMajor { set; get; }
        int NMinor { set; get; }
        double TIncreament { set; get; }
        float FontSize { set; get; }
        double Angle { set; get; }
    }
}
