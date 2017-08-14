using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph
{
    public interface IBarChart
    {
        void CreateBarChart(Mtb.Project proj, Mtb.Worksheet ws,
            BarTypes bType = BarTypes.Stack, BarVOrder barVOrder = BarVOrder.RowOuterMost);
        void SetBarVariable(String inputStr);
        void SetBarDatLabel(bool b);
        void SetLabelVariable(String inputStr);
        void RemoveBarVariable();
        void RemoveLabelVariable();
        //MtbGraphFrame
        void SetExportCommand(bool b, String path);
        void SetCopyToClipboard(bool b);
        void SaveGraph(bool b, String path);
        void SetGraphTitle(String title);
        void SetXAxlabel(String label);
        void SetYAxlabel(String label);
        void SetYScaleMin(double d);
        void SetYScaleMax(double d);
        void SetXAxlabelAngle(double d);
        void SetDefaultXAxlabelAngle();
        void SetDefaultXAxlabel();
        void SetDefaultYAxlabel();
        void SetDefaultTitle();
        void SetDefaultYScale();
        void SetYRefValue(ref double[] values, ref int[] types, ref int[] colors);
        void ClearYRefValue();
        void Dispose();
    }
}
