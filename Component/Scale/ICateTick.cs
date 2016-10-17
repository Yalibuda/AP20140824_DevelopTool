using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Scale
{
    public interface ICateTick
    {
        int Start { set; get; }
        int Increament { set; get; }
        float FontSize { set; get; }
        double Angle { set; get; }
        void SetTShow(dynamic tickindex);
    }
}
