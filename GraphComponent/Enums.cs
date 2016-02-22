using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    public enum TickAttribute
    {
        Position, NumberOfTicks, ByIncrement, Default, ShowAllTickAsPossible
    }

    public enum ScaleType
    {
        X_axis, Y_axis, Secondary_Y_axis
    }

    public enum RefeStatus
    {
        None, Simple, Multi
    }

    public enum LegendboxLook
    {
        Normal, Transparent
    }

    public enum Location
    {
        Auto, RightTop, LeftDown, Custom
    }
    public enum DatlabPlace
    {
        Default, Center, Below
    }

    public enum DatlabType
    {
        Value, RowNum, LabFromColumn
    }

    public enum DatlabColor
    {
        Default, Custom
    }
    public enum LegendNotationType
    {
        Trend, Symbol, Bar
    }

    public enum DataRegionStype
    {
        Default, Transparent
    }

    public enum BarChartType
    {
        Stack, Cluster
    }
    public enum BarChartTableArrangement
    {
        RowsOuterMost, ColsOuterMost
    }
}
