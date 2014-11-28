using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace MtbGraph.GraphComponent
{
    public interface ICOMInterop_Scale_Continuous
    {
        ScaleBoundary Min { get; }
        ScaleBoundary Max { get; }
        ScaleTick Tick { get; }
        AxLabel AxLab { get; }
        Reference Reference { get; }
        ContinuousScale SecsScale { get; }
    }
}
