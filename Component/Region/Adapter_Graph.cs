using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Region
{
    public class Adapter_Graph: IGraph
    {
        public Adapter_Graph(Mtblib.Graph.Component.Region.Region region)
        {
            _region = region;
        }
        Mtblib.Graph.Component.Region.Region _region;

        public void SetSize(double width, double height)
        {
            _region.SetCoordinate(width, height);
        }

        public double[] GetSize()
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
