using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.MyTrend
{
    public interface ICOMInterop_TargetAttribute
    {
        void SetTargetColor(dynamic color);
        dynamic TargetColor { get; }
        void SetTargetType(dynamic linetype);
        dynamic TargetType { get; }
        void SetNotationColor(dynamic color);
        dynamic NotationColor { get; }

    }
}
