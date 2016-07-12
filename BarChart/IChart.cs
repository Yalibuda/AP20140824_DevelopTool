using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.BarChart
{
    public enum ChartFunctionType
    {
        SUM, COUNT, N, NMISS, MEAN, MEDIAN, MINIMUM, MAXIMUM, STDEV, SSQ
    }
    public enum ChartStackType
    {
        Stack, Cluster
    }

    public interface IChart
    {
        void SetVariables(dynamic var);
        void SetGroupingBy(dynamic var);
        void SetPanelBy(dynamic var);
        BarChart.ChartFunctionType FuncType { set; get; }
        BarChart.ChartStackType StackType { set; get; }
        int ColumnAtGroupingLevel { get; set; }
        Component.Scale.ICateScale XScale { get; }
        Component.Scale.IContScale YScale { get; }
        Component.IDatlab Datlab { get; }        
        Component.Region.ILegend Legend { get; }
        Component.ILabel Title { get; }
        Component.IFootnote Footnotes { get; }

        string GSave { set; get; }
        void SetMtbEnvironment(Mtb.Project proj, Mtb.Worksheet ws);
        void Run();
        void Dispose();
        
    }
}
