using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MtbGraph.GraphComponent;

namespace MtbGraph.MyTrend
{
    public interface ICOMInterop_Trend
    {
        Line Line { get; }
        CategoricalScale X_Scale { get; }
        ContinuousScale Y_Scale { get; }
        SimpleLegend LegendBox { get; }
        Annotation Annotation { get; }
        Datlab Datalabel { get; }
        TargetAttribute TargetAttribute { get; }

        void SetVariable(ref Object variables, ScaleType scaletype);
        void SetLabelVariable(ref object variables);
        void SetTargetVariable(ref Object variables, ScaleType scaletype);
        //void SetTargetColor(dynamic colors);
        //void SetTargetType(ref Object linetype);
        void SetGroupVariable(String column);
        void SetMtbEnvironment(Mtb.Project proj, Mtb.Worksheet ws);

        void Run();
        void SaveGraph(bool b, String outputPath);
        void SetExportCommand(bool b, String outputPath = null);
        void CopyGraphToClipboard(bool b);
        void Dispose();


    }
}
