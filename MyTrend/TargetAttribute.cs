using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.MyTrend
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class TargetAttribute : ICOMInterop_TargetAttribute
    {
        public TargetAttribute()
        {

        }
        public dynamic TargetColor { set; get; }
        public dynamic TargetType { set; get; }
        public dynamic NotationColor { set; get; }
        
        public void SetTargetColor(dynamic color)
        {
            this.TargetColor = color;

        }
        public void SetTargetType(dynamic linetype)
        {
            this.TargetType = linetype;

        }
        public void SetNotationColor(dynamic color)
        {
            this.NotationColor = color;
        }

        public TargetAttribute Clone()
        {
            return (TargetAttribute)this.MemberwiseClone();
        }


    }
}
