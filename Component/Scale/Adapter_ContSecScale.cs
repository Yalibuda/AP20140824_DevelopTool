using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Scale
{
    public class Adapter_ContSecScale: IContSecScale
    {
        Mtblib.Graph.Component.Scale.ContSecScale _scale;
        public Adapter_ContSecScale(Mtblib.Graph.Component.Scale.ContSecScale scale)
        {
            _scale = scale;
            _axlab = new Adapter_Lab(_scale.Label);
            _tick = new Adapter_ContTick(_scale.Ticks);
        }

        public dynamic Variables
        {
            get
            {
                return _scale.Variable;
            }
            set
            {
                _scale.Variable = value;
            }
        }

        ILabel _axlab;
        public ILabel AxLab
        {
            set { _axlab = value; }
            get { return _axlab; }
        }

        public double Min
        {
            set { _scale.Min = value; }
            get { return _scale.Min; }
        }

        public double Max
        {
            set { _scale.Max = value; }
            get { return _scale.Max; }
        }

        private IContTick _tick;
        public IContTick Tick
        {
            get
            {
                return _tick;
            }
            set
            {
                _tick = value;
            }
        }
    }
}
