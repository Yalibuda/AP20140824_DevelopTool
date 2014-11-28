using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace MtbGraph.GraphComponent
{
    public interface ICOMInterop_AxLab
    {
        String Label { set; get; }
        void SetDefault();

    }
}
