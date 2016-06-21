using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Scale
{
    public interface ICateScale
    {
        ILabel AxLab { set; get; }
        ICateTick Tick { set; get; }
    }
}
