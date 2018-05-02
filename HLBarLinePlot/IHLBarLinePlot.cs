using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.HLBarLinePlot
{
    public enum BarColorOption
    {
        ByInnerMostGroup,
        Single,
        ByOuterMostGroup
    }
    public interface IHLBarLinePlot
    {
        void SetVariableAtBarChart(dynamic var);
        void SetVariableAtBoxPlot(dynamic var);
        void SetGroupingBy(dynamic var);
        void SetPanelBy(dynamic var);
        void SetDatlabAtBoxPlotIndiv(dynamic var);
        BarChart.ChartFunctionType FuncTypeAtBarChart { set; get; }
        Component.Scale.IContScale YScaleAtBarChart { get; }
        Component.Scale.IContScale YScaleAtBoxPlot { get; }
        Component.Scale.ICateScale XScale { get; }
        Component.Region.IGraph Graph { get; }
        Component.Region.IRegion DataRegionAtBarChart { get; }
        Component.Region.IRegion DataRegionAtBoxPlot { get; }
        Component.ILabel Title { get; }
        Component.IDatlab DatlabAtBarChart { get; }
        Component.IDatlab DatlabAtBoxPlot { get; }
        Component.IDatlab DatlabAtBoxPlotIndiv { get; }
        BarColorOption BarColorType { get; set; }
        DatalabOption DatlabOptionAtBarChart { get; }
        DatalabOption DatlabOptionAtBoxPlot { get; }
        DatalabOption DatlabOptionAtBoxPlotIndiv { get; }

        /// <summary>
        /// 設定 Layout 畫面分割的位置(0~1)，最下方=0
        /// </summary>
        double Division { set; get; }
        string GSave { set; get; }
        void SetMtbEnvironment(Mtb.Project proj, Mtb.Worksheet ws);
        void Run();
        void Dispose();
    }
}
