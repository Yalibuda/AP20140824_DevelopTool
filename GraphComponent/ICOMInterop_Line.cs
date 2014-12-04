using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace MtbGraph.GraphComponent
{
    public interface ICOMInterop_Line
    {
        void SetType(ref Object type);
        //void SetColor(ref Object color); 避免與Trend 繪製衝突先關閉
        void SetSize(ref Object size);
        void SetDefault();
    }
}
