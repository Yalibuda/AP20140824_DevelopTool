using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Scale
{
    public class Adapter_ContScale : IContScale
    {
        Mtblib.Graph.Component.Scale.ContScale _scale;
        /// <summary>
        /// 將元件轉為使用者可使用的介面
        /// </summary>
        /// <param name="contScale">連續型座標軸</param>
        public Adapter_ContScale(Mtblib.Graph.Component.Scale.ContScale contScale)
        {
            _scale = contScale;
            _axlab = new Adapter_Lab(_scale.Label);
            _tick = new Adapter_ContTick(_scale.Ticks);
            _secscale = new Adapter_ContSecScale(_scale.SecScale);
        }

        private ILabel _axlab;
        public ILabel AxLab
        {
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
        }

        private IContSecScale _secscale;
        public IContSecScale SecScale
        {
            get
            {
                return _secscale;
            }
        }

    }
}
