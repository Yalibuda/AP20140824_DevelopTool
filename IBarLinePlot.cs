using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;
using MtbGraph.GraphComponent;

namespace MtbGraph
{

    public interface IBarLinePlot
    {
        Reference BarRef { set; get; }
        Reference TrendRef { set; get; }
        BarLineLegendBox LegendBox { set; get; }
        void CreateBarLinePlot(Mtb.Project proj, Mtb.Worksheet ws, BarTypes btype = BarTypes.Stack);
        void SetBarVariable(String inputStr);
        void SetBarDatLabel(bool b);
        void SetTrendDatLabel(bool b, int decimalNumber = 100);
        void SetTargetDatLabel(bool b);
        void SetLabelVariable(String inputStr);
        void SetScalePrimary(ScalePrimary barScale, ScalePrimary lineScale);
        void SetTrendVariable(String inputStr);
        void SetTargetVariable(String inputStr);
        void SetTypeOfTarget(ref int[] intArr);
        void SetColorOfTarget(ref int[] intArr);
        void RemoveBarVariable();
        void RemoveLabelVariable();
        void RemoveTrendVariable();
        void RemoveTargetVariable();
        void SetDefaultTargetAttribute();


        //MtbGraphFrame
        void SetExportCommand(bool b, String path);
        void SetCopyToClipboard(bool b);
        void SaveGraph(bool b, String path);
        void SetGraphTitle(String title);
        //X-Axis
        void SetXAxlabel(String label);
        void SetXAxlabelAngle(double d);
        void SetDefaultXAxlabel();
        void SetDefaultXAxlabelAngle();
        //Primary scale
        void SetYAxlabel(String label);
        void SetYScaleMin(double d);
        void SetYScaleMax(double d);
        void SetYScaleTick(ScaleTickAttribute atrr, double val);
        void SetDefaultYScale();        
        void SetDefaultYAxlabel();
        //Secondary scale
        void SetSecsAxlabel(String label);        
        void SetSecScaleMin(double d);
        void SetSecScaleMax(double d);
        void SetSecScaleTick(ScaleTickAttribute attr, double val);
        void SetDefaultSecAxlabel();
        void SetDefaultSecScale();
        void SetLegendBoxPosiAutoSetting(bool b);
        void SetSecScaleSize(double d);
        void SetScaleSize(double d);
        void SetSecScaleVisible(bool b);

        //Primary & Secondary scale
        void SetYScaleInt(bool b1, bool b2);
        void SetSameScale(bool b);
        

        void SetDefaultTitle();
        void Dispose();
        
    }
}
