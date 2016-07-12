using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Categoricalplot
{
    public interface IBoxplot
    {
        void SetVariables(dynamic var);
        void SetGroupingBy(dynamic var);
        Component.Scale.ICateScale XScale { get; }
        Component.Scale.IContScale YScale { get; }
        Component.IDatlab Datlab { get; }
        Component.Region.IRegion DataRegion { get; }
        Component.DataView.IDataView Mean { get; }
        Component.DataView.IDataView CMean { get; }
        Component.DataView.IBox RBox { get; }
        Component.DataView.IBox IQRBox { get; }
        Component.DataView.IDataView Individual { get; }
        Component.DataView.IDataView Outlier { get; }
        bool Whisker { set; get; }
        

    }
}
