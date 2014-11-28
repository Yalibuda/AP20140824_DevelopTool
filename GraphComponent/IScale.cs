using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    public interface IScale
    {
        ScaleBoundary Min { set; get; }
        ScaleBoundary Max { set; get; }
        ScaleTick Tick { set; get; }
        AxLabel AxLab { set; get; }
        Reference Reference { set; get; }
        String GetCommand();
        void SetScaleVariable(ref Object varCols, Mtb.Worksheet ws, Mtb.Project proj);

    }
}
