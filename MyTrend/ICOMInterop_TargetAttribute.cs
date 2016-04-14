using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.MyTrend
{
    public interface ICOMInterop_TargetAttribute: GraphComponent.ICOMInterop_Line
    {
        void SetColor(dynamic color);
        void SetType(dynamic linetype);
        void SetSize(dynamic size);
        void SetNotationSize(dynamic fontSize);
        bool ShowNotation { set; get; }
        
        void SetDefault();
        
        
        //dynamic TargetType { get; }
        //void SetNotationColor(dynamic color);
        //dynamic NotationColor { get; }

    }
}
