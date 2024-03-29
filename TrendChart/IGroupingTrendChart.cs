﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.TrendChart
{
    public interface IGroupingTrendChart
    {
        void SetVariables(dynamic var);
        void SetGroupingBy(dynamic var);
        void SetXGroup(dynamic var);
        void SetStamp(dynamic var);
        void SetSymbolSize(dynamic var);
        void SetLineSize(dynamic var);
        void SetGridXVisibile(dynamic var);
        void SetGridYVisibile(dynamic var);
        void SetY1LabelDec(dynamic var);
        void SetY2LabelDec(dynamic var);
        void SetY1LabelVisible(dynamic var);
        void SetY2LabelVisible(dynamic var);
        void SetY1Target(dynamic var);
        void SetY2Target(dynamic var);
        void SetY1LineType(dynamic var);
        void SetY2LineType(dynamic var);
        void SetY1Color(dynamic var);
        void SetY2Color(dynamic var);
        bool IfOnlyLastLabel { get; set; }
        void SetNoOOSDatlabVisible(dynamic var);
        void SetOOSSymbolSize(dynamic var);
        //void SetOOSSymbolColor(dynamic var);
        void SetY1OOSSymbolColor(dynamic var);
        void SetY2OOSSymbolColor(dynamic var);
        void SetY1DatlabColor(dynamic var);
        void SetY2DatlabColor(dynamic var);
        void SetY1OOSDatlabColor(dynamic var);
        void SetY2OOSDatlabColor(dynamic var);
        void SetY1ZeroVisible(dynamic var);
        void SetY2ZeroVisible(dynamic var);
        
        void SetSymbolColor(dynamic var);

        void SetYRefeValues(dynamic var);
        void SetYRefeColors(dynamic var);
        //void SetY2RefeValues(dynamic var);
        //void SetY2RefeColors(dynamic var);

        Component.Scale.IContScale XScale { get; }
        Component.Scale.IContScale YScale { get; }
        Component.IDatlab Datlab { get; }
        Component.Region.ILegend Legend { get; }
        Component.Region.IRegion DataRegion { get; }
        Component.Region.IGraph Graph { get; }
        Component.ILabel Title { get; }
        Component.IFootnote Footnotes { get; }

        string GSave { set; get; }
        void SetMtbEnvironment(Mtb.Project proj, Mtb.Worksheet ws);
        void Run();
        void Dispose();
    }
}
