using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph
{
    public interface ITrendChart
    {
        void CreateTrendChart(Mtb.Project proj, Mtb.Worksheet ws);
        //Trend
        void SetTrendVariable(String inputStr);
        void RemoveTrendVariable();
        void SetTrendDatLabel(bool b, int decimalNumber = 100);
        //Target
        void SetTargetVariable(String inputStr);
        void SetTargetDatLabel(bool b);
        void SetTypeOfTarget(ref int[] intArr);
        void SetColorOfTarget(ref int[] intArr);
        void RemoveTargetVariable();
        void SetDefaultTargetAttribute();
        //Label
        void SetLabelVariable(String inputStr);
        void RemoveLabelVariable();
        //Secondary variable
        void SetSecVariable(String inputStr);
        void RemoveSecVariable();

        //MtbGraphFrame
        void SetExportCommand(bool b, String path);
        void SetCopyToClipboard(bool b);
        void SaveGraph(bool b, String path);
        void SetGraphTitle(String title);
        void SetSecsAxlabel(String label);
        void SetXAxlabel(String label);
        void SetXAxlabelAngle(double d);
        void SetYAxlabel(String label);
        void SetYScaleMin(double d);
        void SetYScaleMax(double d);
        void SetSecScaleMin(double d);
        void SetSecScaleMax(double d);
        void SetDefaultXAxlabel();
        void SetDefaultXAxlabelAngle();
        void SetDefaultYAxlabel();
        void SetDefaultSecAxlabel();
        void SetDefaultTitle();
        void SetDefaultYScale();
        void SetDefaultSecScale();
        void SetYRefValue(ref double[] values, ref int[] types, ref int[] colors);
        void ClearYRefValue();
        void SetSecsRefValue(ref double[] values, ref int[] types, ref int[] colors);
        void ClearSecsRefValue();
        void Dispose();
    }
}
