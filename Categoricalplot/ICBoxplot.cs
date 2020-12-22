using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Categoricalplot
{
    public interface ICBoxplot
    {
        // 要設定的才要開
        bool IfStatVisible { get; set; }
        bool MinVisible { get; set; }
        bool Q1Visible { get; set; }
        bool MedianVisible { get; set; }
        bool MeanVisible { get; set; }
        bool Q3Visible { get; set; }
        bool MaxVisible { get; set; }
        bool CountVisible { get; set; }
        void Run();
        //string DefaultCommand();
        //Component.Scale.ICateScale XScale { get; }
        //Component.Scale.IContScale YScale { get; }
        //Component.IDatlab Datlab { get; }
        //Component.IDatlab DatlabAtBoxPlotIndiv { get; }
        //Component.Region.IRegion DataRegion { get; }
        //Component.DataView.IDataView Mean { get; }
        //Component.DataView.IDataView CMean { get; }
        //////Component.DataView.IBox RBox { get; }
        //Component.DataView.IBox IQRBox { get; }
        //Component.DataView.IDataView Individual { get; }
        //Component.DataView.IDataView Outlier { get; }
        ////bool Whisker { set; get; }


        string GSave { set; get; }
        void SetMtbEnvironment(Mtb.Project proj, Mtb.Worksheet ws);
        void SetVariables(dynamic var);
        void SetGroupingBy(dynamic var);
        //void SetTextSize(dynamic size);
        void Dispose();



    }
}
