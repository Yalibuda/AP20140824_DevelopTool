using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Region
{
    internal class Adapter_Legend: ILegend
    {
        Mtblib.Graph.Component.Region.Legend _legend = null;
        public Adapter_Legend(Mtblib.Graph.Component.Region.Legend legend)
        {
            _legend = legend;
        }
        
        public float FontSize
        {
            get
            {
                return _legend.FontSize;
            }
            set
            {
                _legend.FontSize = value;
            }
        }

        public void SetCoordinate(double xmin, double xmax, double ymin, double ymax)
        {
            _legend.SetCoordinate(xmin, xmax, ymin, ymax);
        }

        public double[] GetCoordinate()
        {
            return _legend.GetCoordinate();
        }

        public bool AutoSize
        {
            get
            {
                return _legend.AutoSize;
            }
            set
            {
                _legend.AutoSize = value;
            }
        }
    }
}
