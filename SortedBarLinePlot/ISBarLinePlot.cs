using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace MtbGraph.SortedBarLinePlot
{
    public interface ISBarLinePlot
    {
        void SetBarVariable(dynamic d);
        void SetTrendVariable(dynamic d);
        void SetGroupingBy(dynamic d);
        //dynamic BarVariable { set; }
        //dynamic TrendVariable { set;  }     
        //dynamic GroupingBy { set;  }
        Component.Scale.IContScale XScale { get; }
        Component.Scale.IContScale YScale { get; }
        Component.IDatlab Datlab { get; }
        Component.ILabel Title { get; }
        Component.Region.IRegion DataRegion { get; }
        string GSave { set; get; }
        void SetMtbEnvironment(Mtb.Project proj, Mtb.Worksheet ws);
        int TopK { set; get; }
        void Run();

    }
}
