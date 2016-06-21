using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Scale
{
    public interface IContSecScale
    {
        dynamic Variables { set; get; }
        ILabel AxLab { set; get; }
        double Min { set; get; }
        double Max { set; get; }
        IContTick Tick { set; get; }
    }
}
