using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace MtbGraph.GraphComponent
{
    public interface ICOMInterop_Tick
    {
        void SetTickPosition(ref Object ticks);
        void SetNumberOfMajorTick(int nmajor);
        int GetNumberOfMajorTick();
        void SetIncrement(double increment);
        double GetIncrement();
        void SetDefault();
        double TickAngle { set; get; }
        double FontSize { set; get; }
    }
}
