using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Region
{
    public interface IGraph
    {
        void SetSize(double width, double height);
        double[] GetSize();
        bool AutoSize { get; set; }
    }
}
