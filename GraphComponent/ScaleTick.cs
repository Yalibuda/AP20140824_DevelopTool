using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class ScaleTick : ICOMInterop_Tick, IScaleTick
    {
        public ScaleTick()
        {
            SetDefault();
        }

        public ScaleTick Clone()
        {           
            /*
             * 因為 Tickposition 複製後也不會修正，所以直接淺層複製就好
             */ 
            return (ScaleTick)this.MemberwiseClone();
        }

        public TickAttribute TickAttr { private set; get; }

        private Object tickposition = null;
        public void SetTickPosition(ref object ticks)
        {
            this.tickposition = ticks;
            this.TickAttr = TickAttribute.Position;
        }

       private int nmajor;
        public void SetNumberOfMajorTick(int nmajor)
        {
            this.nmajor = nmajor;
            this.TickAttr = TickAttribute.NumberOfTicks;
        }

        private double increment;
        public void SetIncrement(double increment = 1.23456E+30)
        {
            if (increment >= 1.23456E+30)
            {
                throw new ArgumentException("Invalid input value of start, end or increment");
                return;
            }
            this.increment = increment;
            this.TickAttr = TickAttribute.ByIncrement;
        }

        public void SetDefault()
        {
            this.nmajor = -1;
            this.increment = 1.23456E+30;
            this.tickposition = null;
            this.TickAttr = TickAttribute.Default;
            this.TickAngle = 1.23456E+30;
            this.FontSize = 8;
        }


        public double TickAngle { get; set; }

        public double FontSize { get;set; }

        public int GetNumberOfMajorTick()
        {
            return this.nmajor;
        }

        public double GetIncrement()
        {
            return this.increment;
        }
    }
}
