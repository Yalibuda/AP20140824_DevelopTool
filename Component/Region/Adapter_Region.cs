﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Region
{
    internal class Adapter_Region : IRegion
    {
        Mtblib.Graph.Component.Region.Region _region;
        public Adapter_Region(Mtblib.Graph.Component.Region.Region region)
        {
            _region = region;
        }

        public void SetCoordinate(double xmin, double xmax, double ymin, double ymax)
        {
            _region.SetCoordinate(xmin, xmax, ymin, ymax);
        }

        public double[] GetCoordinate()
        {
            return _region.GetCoordinate();
        }


        public bool AutoSize
        {
            get
            {
                return _region.AutoSize;
            }
            set
            {
                _region.AutoSize = value;
            }
        }
    }
}
