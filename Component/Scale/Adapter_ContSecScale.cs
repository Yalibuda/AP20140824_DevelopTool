﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Scale
{
    internal class Adapter_ContSecScale : IContSecScale
    {
        private Mtblib.Graph.Component.Scale.ContSecScale _scale;
        public Adapter_ContSecScale(Mtblib.Graph.Component.Scale.ContSecScale scale)
        {
            _scale = scale;
            _axlab = new Adapter_Lab(_scale.Label);
            _tick = new Adapter_ContTick(_scale.Ticks);
            _refe = new Adapter_Refe(_scale.Refes);
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

        public void SetVariables(dynamic d)
        {
            Variables = d;
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

        private IRefe _refe;
        public IRefe Refe
        {
            get { return _refe; }
        }

    }
}
