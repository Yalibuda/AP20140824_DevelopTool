using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MtbGraph.GraphComponent;

namespace MtbGraph.MyBarChart
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class BarchartXScale: CategoricalScale, ICOMInterop_BarchartXScale
    {
         public BarchartXScale(ScaleType scale_axis)
            : base(scale_axis)
        {
            if (scale_axis != ScaleType.X_axis) return;
            this.scale_axis = scale_axis;
        }
    }
}
