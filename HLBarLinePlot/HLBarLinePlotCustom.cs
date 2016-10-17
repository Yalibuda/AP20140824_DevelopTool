using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace MtbGraph.HLBarLinePlot
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class HLBarLinePlotCustom : HLBarLinePlot, IHLBarLinePlotCustom
    {
        public HLBarLinePlotCustom()
            : base()
        {

        }
        public HLBarLinePlotCustom(Mtb.Project proj, Mtb.Worksheet ws)
            : base(proj, ws)
        {
            SetMtbEnvironment(proj, ws);
        }
        //public override dynamic VariablesAtBarChart
        //{
        //    get
        //    {
        //        return base.VariablesAtBarChart;
        //    }
        //    set
        //    {
        //        if (value == null)
        //        {
        //            _varBarchart = null;
        //            return;
        //        }
        //        Mtb.Column[] _cols = Mtblib.Tools.MtbTools.GetMatchColumns(value, _ws);
        //        _varBarchart = _cols;
        //    }
        //}

        public override dynamic VariablesAtBoxPlot
        {
            get
            {
                return base.VariablesAtBoxPlot;
            }
            set
            {
                if (value == null)
                {
                    _varBoxplot = null;
                    return;
                }
                Mtb.Column[] _cols = Mtblib.Tools.MtbTools.GetMatchColumns(value, _ws);
                _varBoxplot = _cols;
            }
        }
        protected override void SetDefault()
        {
            base.SetDefault();
            _boxplot.CMean.Visible = false;
            _boxplot.Mean.Visible = false;
            FuncTypeAtBoxPlot = FuncType.YIELD;
        }
        protected override string GetCommand()
        {
            Mtb.Column[] varBarchart;
            Mtb.Column[] varBoxplot;
            if (VariablesAtBarChart == null || VariablesAtBoxPlot == null)
            {
                throw new ArgumentNullException("HLBarLinePlot 的 Variable 不可為空");
            }
            else
            {
                varBarchart = VariablesAtBarChart;
                varBoxplot = VariablesAtBoxPlot;
            }
            Mtb.Column[] var = varBoxplot.Concat(varBarchart).ToArray();
            // 對應不同的公式，須要有不一樣的檢查機制
            // {...}
            switch (FuncTypeAtBoxPlot)
            {
                case FuncType.PPM:
                case FuncType.YIELD:
                default:
                    if (varBoxplot.Length != 2) throw new ArgumentException(
                           string.Format("Boxplot變數個數={0}，計算 {1} 需要兩個變數。",
                           varBoxplot.Length,
                           FuncTypeAtBoxPlot.ToString()));
                    break;
            }


            Mtb.Column[] gps = GroupingBy;
            Mtb.Column[] pane = PanelBy;

            StringBuilder cmnd = new StringBuilder();
            cmnd.AppendLine("macro");
            cmnd.AppendLine("hlbarline y.1-y.k;"); // 第一個值用於 Bar chart           
            cmnd.AppendLine("group x.1-x.m;"); // 1~m 代表分群由外至內
            cmnd.AppendLine("pane p.");
            cmnd.AppendLine("mcolumn y.1-y.k");
            cmnd.AppendLine("mcolumn x.1-x.m");
            cmnd.AppendLine("mcolumn xx.1-xx.m");
            cmnd.AppendLine("mcolumn txx.1-txx.m");
            cmnd.AppendLine("mcolumn p pp yy xxcord yytext xxconc1 xxconc2");
            cmnd.AppendLine("mconstant ccount");


            if (pane != null && pane[0].GetNumDistinctRows() > 1)
            {
                int distinctRow = pane[0].GetNumDistinctRows();
                cmnd.AppendFormat("mcolumn xcord.1-xcord.{0} ycord.1-ycord.{0} dlab.1-dlab.{0}\r\n", distinctRow);
            }
            else
            {
                cmnd.AppendLine("mcolumn xcord.1 ycord.1 dlab.1");
            }

            cmnd.AppendLine("mcolumn tmpy.1-tmpy.5");

            cmnd.AppendLine("mreset");
            cmnd.AppendLine("notitle");
            cmnd.AppendLine("noecho");
            cmnd.AppendLine("brief 0");

            List<double> vlineInBarChart = new List<double>();

            #region 使用 Minitab command line 取得 BarChart 最外層的垂直線 vline(測試後比 DataTable 法快一點)
            //在262筆資料，3個分群下大概比 DataTable 法快約0.3秒
            if (gps != null && gps.Length > 1)
            {                
                StringBuilder tmpCmnd = new StringBuilder();
                int currentWsCount = _ws.Columns.Count;
                int currentConstCount = _ws.Constants.Count;
                string[] gvalCol = Mtblib.Tools.MtbTools.CreateVariableStrArray(_ws, gps.Length, Mtblib.Tools.MtbVarType.Column);
                tmpCmnd.AppendLine("#取得 BarChart 垂直分割線");
                tmpCmnd.AppendLine("notitle");
                tmpCmnd.AppendLine("brief 0");
                tmpCmnd.AppendFormat("stat {0};\r\n", var[0].SynthesizedName);
                tmpCmnd.AppendFormat("by {0};\r\n", string.Join(" &\r\n", gps.Select(x => x.SynthesizedName)));
                tmpCmnd.AppendLine("noem;");
                tmpCmnd.AppendFormat("gval {0}.\r\n", string.Join(" &\r\n", gvalCol));
                string[] strCol = Mtblib.Tools.MtbTools.CreateVariableStrArray(_ws, 2, Mtblib.Tools.MtbVarType.Column);
                string[] strConst = Mtblib.Tools.MtbTools.CreateVariableStrArray(_ws, 1, Mtblib.Tools.MtbVarType.Constant);
                tmpCmnd.AppendFormat("Count {0} {1}\r\n", gvalCol[0], strConst[0]);
                tmpCmnd.AppendFormat("Set {0}\r\n", strCol[0]);
                tmpCmnd.AppendFormat("1:{0}\r\n", strConst);
                tmpCmnd.AppendLine("End");
                if (gps[0].DataType == Mtb.MtbDataTypes.Text)
                {
                    tmpCmnd.AppendFormat("let {0}=if(Lag({1},1)<>\"\" AND {1}<>Lag({1},1),{0},MISS())\r\n",
                    strCol[0], gvalCol[0]);
                }
                else
                {
                    tmpCmnd.AppendFormat("let {0}=if(Lag({1},1)<>MISS() AND {1}<>Lag({1},1),{0},MISS())\r\n",
                    strCol[0], gvalCol[0]);
                }
                
                tmpCmnd.AppendLine("title");
                tmpCmnd.AppendLine("brief 2");
                string tPath = Mtblib.Tools.MtbTools.BuildTemporaryMacro("~tmpmacro.mtb", tmpCmnd.ToString());
                _proj.ExecuteCommand(string.Format("exec \"{0}\" 1", tPath), _ws);

                double[] rowid = _ws.Columns.Item(strCol[0]).GetData();
                for (int i = _ws.Columns.Count + 1; --i > currentWsCount; ) _ws.Columns.Remove(i);
                for (int i = _ws.Constants.Count + 1; --i > currentConstCount; ) _ws.Constants.Remove(i);
                vlineInBarChart = rowid.Where(x => x < Mtblib.Tools.MtbTools.MISSINGVALUE).Select(x => x - 0.5).ToList();             
            }

            #endregion

            #region 以DataTable 法取得 BarChart 最外層的垂直線 vline(已關閉)
            //if (gps != null && gps.Length > 1)
            //{
            //    #region 計算 BarChart 最外層群組的分割線(垂直輔助線)
            //    Console.WriteLine("{0:mm:ss:fff}\tStart vline calculation", DateTime.Now);
            //    /****************************************************
            //    * 當 grouping by 長度 >=2，最後一個分群做為 cluster 使用，在 HLBarLine 中要全顯示
            //    */
            //    DataTable dt = new DataTable();
            //    Mtb.Column[] gp1;

            //    //如果只有多個 group by，最後一個用於 cluster
            //    gp1 = gps;

            //    // 建立存放 group by 資訊的 datatable，並以 macro 中的分群名稱命名
            //    for (int i = 0; i < gp1.Length; i++)
            //    {
            //        switch (gp1[i].DataType)
            //        {
            //            case Mtb.MtbDataTypes.DataUnassigned:
            //                throw new ArgumentNullException(string.Format("分群欄位{0}無內容", gp1[i].Name));
            //            case Mtb.MtbDataTypes.DateTime:
            //                dt.Columns.Add("xx." + (i + 1), typeof(DateTime));
            //                break;
            //            case Mtb.MtbDataTypes.Numeric:
            //                dt.Columns.Add("xx." + (i + 1), typeof(double));
            //                break;
            //            case Mtb.MtbDataTypes.Text:
            //                dt.Columns.Add("xx." + (i + 1), typeof(string));
            //                break;
            //        }
            //    }
            //    Console.WriteLine("{0:mm:ss:fff}\tTable define Completed", DateTime.Now);
            //    //定義表格大小
            //    for (int i = 0; i < gp1[0].RowCount; i++) dt.Rows.Add(dt.NewRow());
            //    //填入資料
            //    for (int c = 0; c < gp1.Length; c++)
            //    {
            //        dynamic data = gp1[c].GetData();
            //        for (int r = 0; r < gp1[c].RowCount; r++)
            //        {
            //            dt.Rows[r][c] = data[r];
            //        }
            //    }
            //    Console.WriteLine("{0:mm:ss:fff}\tDataFill Completed", DateTime.Now);
            //    //動態對不特定數量的欄位做 Group by
            //    IEnumerable<string> groupCol = gp1.Select((x, i) => "xx." + (i + 1)).ToArray();
            //    var groupedTable = from row in dt.Rows.Cast<DataRow>()
            //                       group row by new Tool.NTuple<object>(from nm in groupCol select row[nm]) into gp
            //                       select gp.Key;
            //    //判斷最外層分群的變化
            //    string[] outGroup = groupedTable.Select(x => x.Values[0].ToString()).OrderBy(x => x).ToArray();
            //    string[] outGroupLag = new string[outGroup.Length]; // 建立 lag1
            //    outGroupLag[0] = null;
            //    Array.Copy(outGroup, 0, outGroupLag, 1, outGroup.Length - 1);
            //    for (int i = 0; i < outGroup.Length; i++)
            //    {
            //        if (outGroup[i] != null && outGroupLag[i] != null)
            //        {
            //            if (outGroup[i] != outGroupLag[i]) vlineInBarChart.Add(0.5 + i);
            //        }
            //    }
            //    Console.WriteLine("{0:mm:ss:fff}\tvline Completed", DateTime.Now);
            //    #endregion
            //} 
            #endregion


            cmnd.AppendLine("layout;");
            if (GSave != null)
            {
                _chart.GraphPath = GSave;
                cmnd.Append(_chart.GetOptionCommand());
                _chart.GraphPath = null;
            }
            cmnd.AppendFormat("title \"{0}({1} & {2})\";\r\n",
                Title.Text == null ? "Bar-Line Plot" : Title.Text,
                 FuncTypeAtBarChart.ToString(),
                FuncTypeAtBoxPlot.ToString());
            cmnd.AppendLine("offset 0 -0.044801;");
            cmnd.AppendLine("graph 10 4;");
            cmnd.AppendLine(".");

            #region 建立 Bar chart
            cmnd.AppendFormat("chart {0}(y.1) &\r\n", _chart.FuncType.ToString());
            if (gps != null)
            {
                cmnd.AppendLine("*x.1;");
                if (gps.Length > 1)
                {
                    cmnd.AppendLine("group x.2-x.m;");
                }
            }
            else
            {
                cmnd.AppendLine(";");
            }

            _chart.XScale = (Mtblib.Graph.Component.Scale.CateScale)_xscale.Clone();
            if (gps != null && gps.Length > 1)
            {
                List<int> tshow = new List<int>();
                for (int i = gps.Length; i >= 1; i--)
                {
                    tshow.Add(i);
                }
                _chart.XScale.Ticks.TShow = tshow.ToArray();
                _chart.Bar.GroupingBy = "x.m";
            }

            if (vlineInBarChart != null && vlineInBarChart.Count > 0)
            {
                _chart.XScale.Refes.Values = vlineInBarChart.ToArray();
                _chart.XScale.Refes.Type = 3;
                _chart.XScale.Refes.Color = 20;
            }
            _chart.XScale.Label.Angle = 0;
            cmnd.Append(_chart.XScale.GetCommand());

            #region DataScale for tick increament 處理程序
            //指定 Tick increament 的處理程序
            if (_chart.YScale.Ticks.Increament < Mtblib.Tools.MtbTools.MISSINGVALUE &&
                _chart.YScale.Ticks.NMajor == -1)
            {
                Func<IEnumerable<double>, double> fun;
                #region 取得函數
                switch (this.FuncTypeAtBarChart)
                {
                    default:
                    case MtbGraph.BarChart.ChartFunctionType.SUM:
                        fun = Mtblib.Tools.Arithmetic.Sum;
                        break;
                    case MtbGraph.BarChart.ChartFunctionType.COUNT:
                        fun = Mtblib.Tools.Arithmetic.Count;
                        break;
                    case MtbGraph.BarChart.ChartFunctionType.N:
                        fun = Mtblib.Tools.Arithmetic.N;
                        break;
                    case MtbGraph.BarChart.ChartFunctionType.NMISS:
                        fun = Mtblib.Tools.Arithmetic.NMiss;
                        break;
                    case MtbGraph.BarChart.ChartFunctionType.MEAN:
                        fun = Mtblib.Tools.Arithmetic.Mean;
                        break;
                    case MtbGraph.BarChart.ChartFunctionType.MEDIAN:
                        fun = Mtblib.Tools.Arithmetic.Median;
                        break;
                    case MtbGraph.BarChart.ChartFunctionType.MINIMUM:
                        fun = Mtblib.Tools.Arithmetic.Min;
                        break;
                    case MtbGraph.BarChart.ChartFunctionType.MAXIMUM:
                        fun = Mtblib.Tools.Arithmetic.Max;
                        break;
                    case MtbGraph.BarChart.ChartFunctionType.STDEV:
                        fun = Mtblib.Tools.Arithmetic.Sum;
                        break;
                    case MtbGraph.BarChart.ChartFunctionType.SSQ:
                        fun = Mtblib.Tools.Arithmetic.Sum;
                        break;
                }
                #endregion

                Mtblib.Tools.GScale barchartScale
                = Mtblib.Tools.MtbTools.GetDataScaleInBarChart(_proj, _ws, varBarchart,
                gps, pane, "", fun);

                string tickString = string.Format("0:{0}/{1}",
                    barchartScale.TMaximum, _chart.YScale.Ticks.Increament);
                _chart.YScale.Ticks.SetTicks(tickString);
            }
            else
            {
                _chart.YScale.Ticks.SetTicks(null);
            }
            #endregion

            cmnd.Append(_chart.YScale.GetCommand());
            cmnd.Append(_chart.Bar.GetCommand());

            if (pane != null)
            {
                _chart.Panel.PaneledBy = "p";
                _chart.Panel.RowColumn = new int[] { 1, pane[0].GetNumDistinctRows() };
                cmnd.Append(_chart.Panel.GetCommand());
            }

            cmnd.Append(_chart.DataLabel.GetCommand());
            cmnd.Append(_chart.Legend.GetCommand());
            cmnd.Append(_chart.FigureRegion.GetCommand());
            cmnd.Append(_chart.Title.GetCommand());
            cmnd.AppendLine("nomiss;");
            cmnd.AppendLine("noem;");
            cmnd.AppendLine("coffset 0.1;");
            cmnd.Append(_chart.GetAnnotationCommand());
            cmnd.AppendLine(".");
            #endregion

            #region 計算客製的平均位置
            /*
             * 處理方式會因是否有 Panel Data 有很大的差異；有 Panel data 時，
             * 不同 Panel 的點、線、文字對應不同的 Unit
             * 
             * 
             */
            #region 將分組資訊轉換成數值，用於最後繪製 Marker, Line, Text
            if (gps != null)
            {
                cmnd.AppendLine("stat y.1;");
                cmnd.AppendLine("by x.1-x.m;");
                cmnd.AppendLine("noem;");
                cmnd.AppendLine("gval xx.1-xx.m.");
                cmnd.AppendLine("Count xx.1 ccount");
                cmnd.AppendLine("Text xx.1-xx.m txx.1-txx.m");
                if (gps.Length > 1)
                {
                    cmnd.AppendLine("conc txx.1-txx.m xxconc1");
                }
                else
                {
                    cmnd.AppendLine("copy txx.1 xxconc1");
                }

            }
            else
            {
                cmnd.AppendLine("copy 1 ccount");
            }


            cmnd.AppendLine("set xxcord");
            cmnd.AppendLine(" 1:ccount");
            cmnd.AppendLine(" end"); // 到此為止建立一了一個 conversion table: xxconc1 xxcord
            #endregion

            #region 計算要標示的值
            switch (FuncTypeAtBoxPlot)
            {
                case FuncType.PPM:
                    cmnd.AppendLine("let yy = y.2");
                    cmnd.AppendLine("stat yy;");
                    if (pane != null || gps != null) cmnd.AppendLine("by &");
                    if (pane != null) cmnd.AppendLine(" p &");
                    if (gps != null) cmnd.AppendLine(" x.1-x.m;");
                    cmnd.AppendLine("noem;");
                    cmnd.AppendLine("sums tmpy.1;");
                    if (pane != null || gps != null) cmnd.AppendLine("gval &");
                    if (pane != null) cmnd.AppendLine(" pp &");
                    if (gps != null) cmnd.AppendLine(" xx.1-xx.m;");
                    cmnd.AppendLine(".");

                    cmnd.AppendLine("stat y.3;");
                    if (pane != null || gps != null) cmnd.AppendLine("by &");
                    if (pane != null) cmnd.AppendLine(" p &");
                    if (gps != null) cmnd.AppendLine(" x.1-x.m;");
                    cmnd.AppendLine("noem;");
                    cmnd.AppendLine("sums tmpy.2.");
                    cmnd.AppendLine("let yy = tmpy.1/tmpy.2*(10^6)");
                    cmnd.AppendLine("fnum yy;");
                    cmnd.AppendLine("fixed 0.");
                    cmnd.AppendLine("let tmpy.1 = y.2/y.3*(10^6)");
                    break;
                case FuncType.YIELD:
                default:
                    cmnd.AppendLine("let yy = y.2");
                    cmnd.AppendLine("stat yy;");
                    if (pane != null || gps != null) cmnd.AppendLine("by &");
                    if (pane != null) cmnd.AppendLine(" p &");
                    if (gps != null) cmnd.AppendLine(" x.1-x.m;");
                    cmnd.AppendLine("noem;");
                    cmnd.AppendLine("sums tmpy.1;");
                    if (pane != null || gps != null) cmnd.AppendLine("gval &");
                    if (pane != null) cmnd.AppendLine(" pp &");
                    if (gps != null) cmnd.AppendLine(" xx.1-xx.m;");
                    cmnd.AppendLine(".");

                    cmnd.AppendLine("stat y.3;");
                    if (pane != null || gps != null) cmnd.AppendLine("by &");
                    if (pane != null) cmnd.AppendLine(" p &");
                    if (gps != null) cmnd.AppendLine(" x.1-x.m;");
                    cmnd.AppendLine("noem;");
                    cmnd.AppendLine("sums tmpy.2.");
                    cmnd.AppendLine("let yy = tmpy.1/tmpy.2*100");
                    cmnd.AppendLine("fnum yy;");
                    cmnd.AppendLine("fixed 2.");
                    cmnd.AppendLine("let tmpy.1 = y.2/y.3*100");
                    break;
            }

            #endregion

            #region 建立 X coordinate 規則 (類別 --> 圖上的位置)
            if (gps != null)
            {
                cmnd.AppendLine("text xx.1-xx.m txx.1-txx.m");
                if (gps.Length > 1)
                {
                    cmnd.AppendLine("conc txx.1-txx.m xxconc2");
                }
                else
                {
                    cmnd.AppendLine("copy txx.1 xxconc2");
                }

                cmnd.AppendLine("conv xxconc1 xxcord xxconc2 xxcord");
            }
            else
            {
                cmnd.AppendLine("count yy ccount");
                cmnd.AppendLine("set xxcord");
                cmnd.AppendLine(" 1(1)ccount");
                cmnd.AppendLine(" end");
            }
            #endregion

            #region 建立 Datlab 的值與座標
            if (pane != null && pane[0].GetNumDistinctRows() > 1) //根據 Panel 資料分群
            {
                int distinctRow = pane[0].GetNumDistinctRows();
                cmnd.AppendLine("unstack (yy xxcord) &");
                for (int i = 1; i <= distinctRow; i++)
                {
                    cmnd.AppendFormat("(ycord.{0} xcord.{0}) &\r\n", i);
                }
                cmnd.AppendLine(";");
                cmnd.AppendLine("subs pp.");
                //建立標籤                                
                cmnd.AppendFormat("text ycord.1-ycord.{0} dlab.1-dlab.{0}\r\n", distinctRow);
            }
            else
            {
                cmnd.AppendLine("copy xxcord xcord.1");
                cmnd.AppendLine("copy yy ycord.1");
                cmnd.AppendLine("text yy dlab.1");
            }
            #endregion

            #endregion

            #region 建立 Boxplot
            cmnd.AppendLine("boxplot tmpy.1 &");
            if (gps != null)
            {
                cmnd.AppendLine("*x.1;");
                if (gps.Length > 1)
                {
                    cmnd.AppendLine("group x.2-x.m;");
                }
            }
            else
            {
                cmnd.AppendLine(";");
            }
            _boxplot.XScale = (Mtblib.Graph.Component.Scale.CateScale)_xscale.Clone();
            _boxplot.XScale.Ticks.HideAllTick = true;
            _boxplot.XScale.Label.Visible = false;
            cmnd.Append(_boxplot.XScale.GetCommand());
            if (_boxplot.YScale.Label.Text == null) _boxplot.YScale.Label.Text = FuncTypeAtBoxPlot.ToString();

            #region DataScale for tick increament 處理程序
            //指定 Tick increament 的處理程序
            if (_boxplot.YScale.Ticks.Increament < Mtblib.Tools.MtbTools.MISSINGVALUE &&
                _boxplot.YScale.Ticks.NMajor == -1)
            {
                Mtblib.Tools.GScale boxplotScale
                = Mtblib.Tools.MtbTools.GetDataScaleInCateChart(
                _proj, _ws, varBoxplot, gps, pane, "", true);

                string tickString = string.Format("{0}:{1}/{2}",
                    boxplotScale.TMinimum, boxplotScale.TMaximum, _boxplot.YScale.Ticks.Increament);
                _boxplot.YScale.Ticks.SetTicks(tickString);
            }
            else
            {
                _boxplot.YScale.Ticks.SetTicks(null);
            }
            #endregion

            cmnd.Append(_boxplot.YScale.GetCommand());
            cmnd.AppendLine("nojitter;");
            cmnd.AppendLine("nomiss;");
            cmnd.AppendLine("noem;");
            cmnd.AppendLine("coffset 0.1;");
            cmnd.Append(_boxplot.Individual.GetCommand());
            cmnd.Append(_boxplot.IQRBox.GetCommand());
            cmnd.Append(_boxplot.RBox.GetCommand());
            cmnd.Append(_boxplot.Whisker.GetCommand());
            cmnd.Append(_boxplot.Outlier.GetCommand());

            if (pane != null)
            {
                _boxplot.Panel.PaneledBy = "p";
                _boxplot.Panel.RowColumn = new int[] { 1, pane[0].GetNumDistinctRows() };
                cmnd.Append(_boxplot.Panel.GetCommand());
            }

            cmnd.Append(_boxplot.FigureRegion.GetCommand());
            cmnd.Append(_boxplot.Title.GetCommand());

            if (pane != null && pane[0].GetNumDistinctRows() > 1)
            {
                #region 貼上每一個 Panel 對應的 Customed Mean symbol, Conn 和 Datlab
                for (int i = 0; i < pane[0].GetNumDistinctRows(); i++)
                {
                    string xcord = string.Format("xcord.{0}", i + 1);
                    string ycord = string.Format("ycord.{0}", i + 1);
                    string dlab = string.Format("dlab.{0}", i + 1);
                    Mtblib.Graph.Component.Annotation.Line line = new Mtblib.Graph.Component.Annotation.Line()
                    {
                        Size = 1,
                        Type = 1,
                        Color = 64,
                        Unit = i + 1
                    };
                    line.SetCoordinate(xcord, ycord);
                    _boxplot.ALineLst.Add(line);
                    Mtblib.Graph.Component.Annotation.Marker mark = new Mtblib.Graph.Component.Annotation.Marker()
                    {
                        Size = 1.5,
                        Type = 6,
                        Color = 64,
                        Unit = i + 1
                    };
                    mark.SetCoordinate(xcord, ycord);

                    _boxplot.AMarkerLst.Add(mark);
                    if (_boxplot.MeanDatlab.Visible)
                    {
                        Mtblib.Graph.Component.Annotation.Textbox tbox = new Mtblib.Graph.Component.Annotation.Textbox();
                        tbox.SetCoordinate(xcord, ycord);
                        tbox.Text = dlab;
                        tbox.Unit = i + 1;
                        if (_boxplot.MeanDatlab.Offset != null)
                            tbox.Offset = new double[] { 0.01, 0 };
                        else
                            tbox.Offset = _boxplot.MeanDatlab.Offset;

                        if (_boxplot.MeanDatlab.Placement != null)
                            tbox.Placement = _boxplot.MeanDatlab.Placement;
                        else
                            tbox.Offset = new double[] { 1, 0 };


                        tbox.FontColor = _boxplot.MeanDatlab.FontColor;
                        tbox.FontSize = _boxplot.MeanDatlab.FontSize;
                        tbox.Bold = _boxplot.MeanDatlab.Bold;
                        tbox.Italic = _boxplot.MeanDatlab.Italic;
                        tbox.Underline = _boxplot.MeanDatlab.Underline;
                        _boxplot.ATextLst.Add(tbox);
                    }

                }
                #endregion
            }
            else
            {
                #region 貼上 Customed Mean Symbol, Conn 和 Datlab
                string xcord = "xcord.1";
                string ycord = "ycord.1";
                string dlab = "dlab.1";
                Mtblib.Graph.Component.Annotation.Line line = new Mtblib.Graph.Component.Annotation.Line()
                {
                    Size = 1,
                    Type = 1,
                    Color = 64
                };
                line.SetCoordinate(xcord, ycord);
                _boxplot.ALineLst.Add(line);
                Mtblib.Graph.Component.Annotation.Marker mark = new Mtblib.Graph.Component.Annotation.Marker()
                {
                    Size = 1.5,
                    Type = 6,
                    Color = 64
                };
                mark.SetCoordinate(xcord, ycord);

                _boxplot.AMarkerLst.Add(mark);
                if (_boxplot.MeanDatlab.Visible)
                {
                    Mtblib.Graph.Component.Annotation.Textbox tbox = new Mtblib.Graph.Component.Annotation.Textbox();
                    tbox.SetCoordinate(xcord, ycord);
                    tbox.Text = dlab;
                    if (_boxplot.MeanDatlab.Offset != null)
                        tbox.Offset = _boxplot.MeanDatlab.Offset;
                    else
                        tbox.Offset = new double[] { 0.01, 0 };

                    if (_boxplot.MeanDatlab.Placement != null)
                        tbox.Placement = _boxplot.MeanDatlab.Placement;
                    else
                        tbox.Placement = new double[] { 1, 0 };

                    tbox.FontColor = _boxplot.MeanDatlab.FontColor;
                    tbox.FontSize = _boxplot.MeanDatlab.FontSize;
                    tbox.Bold = _boxplot.MeanDatlab.Bold;
                    tbox.Italic = _boxplot.MeanDatlab.Italic;
                    tbox.Underline = _boxplot.MeanDatlab.Underline;
                    _boxplot.ATextLst.Add(tbox);
                }
                #endregion
            }
            cmnd.Append(_boxplot.GetAnnotationCommand());
            cmnd.AppendLine(".");
            #endregion

            cmnd.AppendLine("endlayout");

            cmnd.AppendLine("endmacro");

            return cmnd.ToString();
        }

        public FuncType FuncTypeAtBoxPlot { set; get; }
    }
    /// <summary>
    /// Special High Level Bar-Line Plot 要 Datlab 的使用公式
    /// </summary>
    public enum FuncType
    {
        PPM, YIELD
    }
}
