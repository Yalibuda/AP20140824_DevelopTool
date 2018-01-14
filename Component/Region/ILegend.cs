using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Region
{
    public interface ILegend
    {
        float FontSize { set; get; }
        void SetCoordinate(double xmin, double xmax, double ymin, double ymax);
        double[] GetCoordinate();
        bool AutoSize { get; set; }
        bool HideLegend { get; set; }
    }
}
