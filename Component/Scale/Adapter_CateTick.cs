using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Scale
{
    internal class Adapter_CateTick : ICateTick
    {
        Mtblib.Graph.Component.Scale.Tick _tick;
        public Adapter_CateTick(Mtblib.Graph.Component.Scale.Tick tick)
        {
            _tick = tick;
        }

        public int Start
        {
            get
            {
                return _tick.Start;
            }
            set
            {
                _tick.Start = value;
            }
        }

        public int Increament
        {
            get
            {
                return (int)_tick.Increament;
            }
            set
            {
                _tick.Increament = value;
            }
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
        public double Angle
        {
            get { return _tick.Angle; }
            set { _tick.Angle = value; }
        }

        public void SetTShow(dynamic tickindex)
        {
            _tick.TShow = tickindex;
        }
    }
}
