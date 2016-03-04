using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MtbGraph.GraphComponent;

namespace MtbGraph.MyBarChart
{
    public interface ICOMInterop_Barchart
    {
        BarchartXScale X_Scale { get; }
        BarchartYScale Y_Scale { get; }
        SimpleLegend LegendBox { get; }
        Annotation Annotation { get; }
        Datlab Datalabel { get; }

        /// <summary>
        /// 設定 Bar Chart 的排列方式
        /// </summary>
        BarChartTableArrangement TableArrangement { set; get; }
        BarChartType ChartType { set; get; }

        void SetVariable(ref Object variables);
        void SetLabelVariable(ref object variables);   
        void SetMtbEnvironment(Mtb.Project proj, Mtb.Worksheet ws);

        void SetDatalabelColor(int color);

        void Run();
        void SaveGraph(bool b, String outputPath);
        void SetExportCommand(bool b, String outputPath = null);
        void CopyGraphToClipboard(bool b);
        void Dispose();
    }
}
