using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.TrendChart
{
    public interface IGroupingTrendChart
    {
        void SetVariables(dynamic var);
        void SetGroupingBy(dynamic var);
        void SetXGroup(dynamic var);
        void SetStamp(dynamic var);
        Component.Scale.IContScale XScale { get; }
        Component.Scale.IContScale YScale { get; }
        Component.IDatlab Datlab { get; }
        Component.Region.ILegend Legend { get; }
        Component.Region.IRegion DataRegion { get; }
        Component.Region.IGraph Graph { get; }
        Component.ILabel Title { get; }
        Component.IFootnote Footnotes { get; }        

        string GSave { set; get; }
        void SetMtbEnvironment(Mtb.Project proj, Mtb.Worksheet ws);
        void Run();
        void Dispose();
    }
}
