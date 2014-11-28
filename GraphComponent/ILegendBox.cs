using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    public interface ILegendBox
    {
        Size GetSize();
        bool Show { set; get; }
        LegendboxLook LegendBoxLook { set; get; }
        LegendNotationType NotationType { set; get; }
        Location Location { set; get; }
        String SectTitle { set; get; }
        bool HideHead { set; get; }
        void SetVariables(ref String[] variables);
        Size CalculateSize(String[] variables);
        String GetCommand();
        void SetDefault();

    }
}
