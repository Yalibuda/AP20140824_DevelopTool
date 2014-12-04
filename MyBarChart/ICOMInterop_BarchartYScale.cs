using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MtbGraph.GraphComponent;

namespace MtbGraph.MyBarChart
{
    public interface ICOMInterop_BarchartYScale
    {
        ScaleBoundary Min { get; }
        ScaleBoundary Max { get; }
        ScaleTick Tick { get; }
        AxLabel AxLab { get; }
        Reference Reference { get; }
        
    }
}
