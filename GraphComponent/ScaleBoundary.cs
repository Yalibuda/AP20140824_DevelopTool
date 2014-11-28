using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    [ClassInterface(ClassInterfaceType.AutoDual)]
    public class ScaleBoundary
    {
        public double Value { set; get; }
        public ScaleBoundary()
        {
            this.Value = 1.23456E+30;
        }
        public void SetDefault()
        {
            this.Value = 1.23456E+30;
        }
        public ScaleBoundary Clone()
        {
            return (ScaleBoundary)this.MemberwiseClone();
        }
    }

}
