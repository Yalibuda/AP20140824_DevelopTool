using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Scale
{
    public interface IContScale
    {
        ILabel AxLab { get; }
        double Min { set; get; }
        double Max { set; get; }
        IContTick Tick { get; }
        IRefe Refe { get; }
        IContSecScale SecScale { get; }

    }
}
