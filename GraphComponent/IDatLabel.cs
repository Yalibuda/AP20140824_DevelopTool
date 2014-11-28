using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    public interface IDatLabel
    {
        bool Show { set; get; }
        DatlabType LabelType { set; get; }
        DatlabColor Color { set; get; }
        void SetDatlabInvisible(List<DatlabModelAttribute> modelAttribute);
        void SetCustomDatlab(List<DatlabModelAttribute> modelAttribute);
        void SetLabelFromColumn(String col);
        String GetCommand();
        void SetDefault();
    }
}
