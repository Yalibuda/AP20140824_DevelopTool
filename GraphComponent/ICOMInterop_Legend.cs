using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace MtbGraph.GraphComponent
{
    public interface ICOMInterop_Legend
    {
        bool Show { set; get; }
        LegendboxLook LegendBoxLook { set; get; }       
        bool HideHead { set; get; }
        String SectTitle { set; get; }
        void SetDefault();

    }
}
