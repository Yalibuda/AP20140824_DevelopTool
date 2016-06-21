using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Scale
{
    public class Adapter_ContTick : IContTick
    {
        Mtblib.Graph.Component.Scale.Tick _tick;
        public Adapter_ContTick(Mtblib.Graph.Component.Scale.Tick tick)
        {
            _tick = tick;

        }
        public int NMajor
        {
            set { _tick.NMajor = value; }
            get { return _tick.NMajor; }
        }
        public int NMinor
        {
            set { _tick.NMinor = value; }
            get { return _tick.NMinor; }
        }
        public float FontSize
        {
            get
            {
                return _tick.FontSize;
            }
            set
            {
                _tick.FontSize = value;
            }
        }

    }
}
