using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Scale
{
    internal class Adapter_CateScale : ICateScale
    {
        Mtblib.Graph.Component.Scale.CateScale _scale;        
        public Adapter_CateScale(Mtblib.Graph.Component.Scale.CateScale scale)
        {
            _scale = scale;
            _axlab = new Adapter_Lab(_scale.Label);
            _tick = new Adapter_CateTick(_scale.Ticks);
            _refe = new Adapter_Refe(_scale.Refes);
        }

        private ILabel _axlab;
        public ILabel AxLab
        {
            set { _axlab = value; }
            get { return _axlab; }
        }

        private ICateTick _tick;
        public ICateTick Tick
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

        private IRefe _refe;
        public IRefe Refe
        {
            get { return _refe; }
        }
    }
}
