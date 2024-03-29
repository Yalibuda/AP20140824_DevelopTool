﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;

namespace MtbGraph.HLBarLinePlot
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class HLBarLinePlot : IHLBarLinePlot, IDisposable
    {
        protected Mtb.Column[] _maxvar = null;
        public virtual dynamic Maxvar
        {
            set
            {
                if (value == null)
                {
                    _maxvar = null;
                    return;
                }
                Mtb.Column[] _cols = Mtblib.Tools.MtbTools.GetMatchColumns(value, _ws);
                if (_cols.Length > 1)
                    throw new ArgumentException(
                        string.Format("輸入欄位數 {0}，不是合法的數量。HLBarLinePlot 只能繪製一個欄位。", _cols.Length));
                _maxvar = _cols;
            }
            get
            {
                return _maxvar;
            }
        }

        /// <summary>
        /// 輸入單一 Variable 時，可以對數據套用的函數
        /// </summary>
        public enum ChartFunctionType
        {
            SUM, COUNT, N, NMISS, MEAN, MEDIAN, MINIMUM, MAXIMUM, STDEV, SSQ
        }

        public HLBarLinePlot()
        {

        }

        public HLBarLinePlot(Mtb.Project proj, Mtb.Worksheet ws)
        {
            SetMtbEnvironment(proj, ws);
        }

        protected Mtb.Project _proj;
        protected Mtb.Worksheet _ws;
        protected Mtblib.Graph.BarChart.Chart _chart;
        protected Mtblib.Graph.CategoricalChart.BoxPlot _boxplot;

        protected Mtb.Column[] _varBarchart = null;
        public virtual dynamic VariablesAtBarChart
        {
            set
            {
                if (value == null)
                {
                    _varBarchart = null;
                    return;
                }
                Mtb.Column[] _cols = Mtblib.Tools.MtbTools.GetMatchColumns(value, _ws);
                if (_cols.Length > 1)
                    throw new ArgumentException(
                        string.Format("輸入欄位數 {0}，不是合法的數量。HLBarLinePlot 只能繪製一個欄位。", _cols.Length));
                _varBarchart = _cols;
            }
            get
            {
                return _varBarchart;
            }
        }
        /// <summary>
        /// 設定要繪製BarChart的資料欄，合法的輸入為單一欄位資訊(string)
        /// </summary>
        /// <param name="var"></param>
        public void SetVariableAtBarChart(dynamic var)
        {
            VariablesAtBarChart = var;
        }


        protected Mtb.Column[] _varBoxplot = null;
        public virtual dynamic VariablesAtBoxPlot
        {
            set
            {
                if (value == null)
                {
                    _varBoxplot = null;
                    return;
                }
                Mtb.Column[] _cols = Mtblib.Tools.MtbTools.GetMatchColumns(value, _ws);
                if (_cols.Length > 1)
                    throw new ArgumentException(
                        string.Format("輸入欄位數 {0}，不是合法的數量。HLBarLinePlot 只能繪製一個欄位。", _cols.Length));
                _varBoxplot = _cols;
            }
            get
            {
                return _varBoxplot;
            }
        }
        /// <summary>
        /// 設定要繪製BoxPlot的資料欄，合法的輸入為單一欄位資訊(string)
        /// </summary>
        /// <param name="var"></param>
        public void SetVariableAtBoxPlot(dynamic var)
        {
            VariablesAtBoxPlot = var;
        }


        protected Mtb.Column[] _groupBy = null;
        public virtual dynamic GroupingBy
        {
            set
            {
                if (value == null)
                {
                    _groupBy = null;
                    return;
                }
                Mtb.Column[] _cols = Mtblib.Tools.MtbTools.GetMatchColumns(value, _ws);
                if (_cols.Length > 3)
                    throw new ArgumentException(
                        string.Format("輸入欄位數 {0}，不是合法的數量。HLBarLinePlot 用於分群的欄位數量最多為 3。", _cols.Length));
                _groupBy = _cols;
            }
            get
            {
                return _groupBy;
            }
        }
        /// <summary>
        /// 設定要分群的資料欄，合法的輸入為單一(string)或多個(string[], 最多3個)欄位
        /// </summary>
        /// <param name="var"></param>
        public void SetGroupingBy(dynamic var)
        {
            GroupingBy = var;
        }


        protected Mtb.Column[] _pane = null;
        public virtual dynamic PanelBy
        {
            set
            {
                if (value == null)
                {
                    _pane = null;
                    return;
                }
                Mtb.Column[] _cols = Mtblib.Tools.MtbTools.GetMatchColumns(value, _ws);
                if (_cols.Length > 1)
                    throw new ArgumentException(
                        string.Format("輸入欄位數 {0}，不是合法的數量。HLBarLinePlot 只能繪製一個欄位。", _cols.Length));
                _pane = _cols;
            }
            get
            {
                return _pane;
            }
        }
        /// <summary>
        /// 設定用於分割畫面的欄位，合法的輸入為單一欄位資訊(string)
        /// </summary>
        /// <param name="var"></param>
        public void SetPanelBy(dynamic var)
        {
            PanelBy = var;
        }

        /// <summary>
        /// 指定或取得要使用於 BarChart 的函數類型
        /// </summary>
        public BarChart.ChartFunctionType FuncTypeAtBarChart
        {
            set
            {
                _chart.FuncType = (Mtblib.Graph.BarChart.Chart.ChartFunctionType)Enum.Parse(
                    typeof(Mtblib.Graph.BarChart.Chart.ChartFunctionType), value.ToString());
            }
            get
            {
                return (BarChart.ChartFunctionType)Enum.Parse(
                    typeof(BarChart.ChartFunctionType), _chart.FuncType.ToString());
            }
        }

        /// <summary>
        /// 取得 Bar Chart 的 Y 軸元件
        /// </summary>
        public Component.Scale.IContScale YScaleAtBarChart
        {
            get
            {
                return new Component.Scale.Adapter_ContScale(_chart.YScale);
            }
        }


        /// <summary>
        /// 取得 Box Plot 的 Y 軸元件
        /// </summary>
        public Component.Scale.IContScale YScaleAtBoxPlot
        {
            get
            {
                return new Component.Scale.Adapter_ContScale(_boxplot.YScale);
            }
        }


        protected Mtblib.Graph.Component.Scale.CateScale _xscale;
        /// <summary>
        /// 取得圖形中的 X 軸元件
        /// </summary>
        public Component.Scale.ICateScale XScale
        {
            get { return new Component.Scale.Adapter_CateScale(_xscale); }
        }

        /// <summary>
        /// 取得圖形中的 Graph region 元件
        /// </summary>
        public Component.Region.IGraph Graph
        {
            get
            {
                //因為沒有定義 Layout 時候的 Graph 元件，所以用 Chart 的 Graph 元件代替，實作程式碼的時候，建立 Graph 指令。
                return new Component.Region.Adapter_Graph(_chart.GraphRegion);
            }
        }

        /// <summary>
        /// 取得圖形中 Bar chart 的 Data region 元件
        /// </summary>
        public Component.Region.IRegion DataRegionAtBarChart
        {
            get { return new Component.Region.Adapter_Region(_chart.DataRegion); }
        }

        /// <summary>
        /// 取得圖形中 Box plot 的 Data region 元件
        /// </summary>
        public Component.Region.IRegion DataRegionAtBoxPlot
        {
            get { return new Component.Region.Adapter_Region(_boxplot.DataRegion); }
        }

        private Mtblib.Graph.Component.Title _title;
        /// <summary>
        /// 取得主標題元件
        /// </summary>
        public Component.ILabel Title
        {
            get { return new Component.Adapter_Lab(_title); }
        }

        /// <summary>
        /// 取得 Bar Chart Datalab 的物件
        /// </summary>
        public Component.IDatlab DatlabAtBarChart
        {
            get
            {
                return new Component.Adapter_DatLab(_chart.DataLabel);
            }
        }

        /// <summary>
        /// 取得 Boxplot Chart Datalab 的物件
        /// </summary>
        public Component.IDatlab DatlabAtBoxPlot
        {
            get
            {
                return new Component.Adapter_DatLab(_boxplot.MeanDatlab);
            }
        }

        /// <summary>
        /// 取得 Boxplot Chart Individual Datalab 的物件
        /// </summary>
        public Component.IDatlab DatlabAtBoxPlotIndiv
        {
            get
            {
                return new Component.Adapter_DatLab(_boxplot.IndivDatlab);
            }
        }

        /// <summary>
        /// 設定 Boxplot 中 Individual Datlab 要顯示的方式
        /// </summary>
        public void SetDatlabAtBoxPlotIndiv(dynamic type)
        {
            string typestr = string.Format("{0}", type);
            switch (typestr)
            {
                case "All":
                    _boxplot.IndivDatlab.DatlabType = Mtblib.Graph.Component.Datlab.DisplayType.YValue;
                    break;
                case "MaxMinOnly":
                    _boxplot.IndivDatlab.DatlabType = Mtblib.Graph.Component.Datlab.DisplayType.Column;
                    break;
                case "MaxOnly":
                    _boxplot.IndivDatlab.DatlabType = Mtblib.Graph.Component.Datlab.DisplayType.Column;
                    _minVisible = false;
                    break;
                case "MinOnly":
                    _boxplot.IndivDatlab.DatlabType = Mtblib.Graph.Component.Datlab.DisplayType.Column;
                    _maxVisible = false;
                    break;
                default:
                    break;
            }
        }
        
        protected BarColorOption _barColorOpt = BarColorOption.Single;
        /// <summary>
        /// 設定或取得 BarChart 分群的顏色設定
        /// </summary>
        public BarColorOption BarColorType
        {
            get
            {
                return _barColorOpt;
            }
            set
            {
                _barColorOpt = value;
            }
        }

        /// <summary>
        /// 設定BarChart 中 Datlab 要顯示的位數
        /// </summary>
        public DatalabOption DatlabOptionAtBarChart { get { return _datlabOptAtChart; } }
        private DatalabOption _datlabOptAtChart;

        /// <summary>
        /// 設定Boxplot 中 Datlab 要顯示的位數
        /// </summary>
        public DatalabOption DatlabOptionAtBoxPlot { get { return _datlabOptAtBoxPlot; } }
        private DatalabOption _datlabOptAtBoxPlot;

        /// <summary>
        /// 設定Boxplot 中 Individual Datlab 要顯示的位數
        /// </summary>
        public DatalabOption DatlabOptionAtBoxPlotIndiv { get { return _datlabOptAtBoxPlotIndiv; } }
        private DatalabOption _datlabOptAtBoxPlotIndiv;

        private double _boxplotmeansymbolsize;
        /// <summary>
        /// 設定 Mean symbol size
        /// </summary>
        public double BoxPlotMeanSymbolSize {
            //get { return _boxplot.Mean.Size; }
            get { return _boxplotmeansymbolsize; }
            set
            {
                _boxplotmeansymbolsize = value;
                _boxplot.Mean.Size = (dynamic)_boxplotmeansymbolsize;
            }
        }

        private double _boxplotindivsymbolsize;
        /// <summary>
        /// 設定 Individual symbol size
        /// </summary>
        public double BoxPlotIndivSymbolSize
        {
            get { return _boxplotindivsymbolsize; }
            //get { return _boxplot.Individual.Size; }
            set
            {
                _boxplotindivsymbolsize = value;
                _boxplot.Individual.Size = (dynamic)_boxplotindivsymbolsize;
            }
        }


        private double _boxPlotMeanConnectLineSize;
        public double BoxPlotMeanConnectLineSize 
        { 
            get 
            {
                return _boxplot.CMean.Size;
            }
            set
            {
                _boxPlotMeanConnectLineSize = value;
                _boxplot.CMean.Size = (dynamic)_boxPlotMeanConnectLineSize;
            }
        }

        /// <summary>
        /// 設定分割的比例值(由下往上所占的比例)
        /// </summary>
        public double Division
        {
            set
            {
                _division = value;
                _chart.FigureRegion.SetCoordinate(0, 1, 0, value);
                _boxplot.FigureRegion.SetCoordinate(0, 1, value, 1);
            }
            get
            {
                return _division;
            }
        }
        private double _division;


        public string GSave { set; get; }

        public void SetMtbEnvironment(Mtb.Project proj, Mtb.Worksheet ws)
        {
            _proj = proj;
            _ws = ws;
            _xscale = new Mtblib.Graph.Component.Scale.CateScale(Mtblib.Graph.Component.ScaleDirection.X_Axis);
            _chart = new Mtblib.Graph.BarChart.Chart(_proj, _ws);
            _boxplot = new Mtblib.Graph.CategoricalChart.BoxPlot(_proj, _ws);
            _datlabOptAtChart = new DatalabOption();
            _datlabOptAtBoxPlot = new DatalabOption();
            _datlabOptAtBoxPlotIndiv = new DatalabOption();
            SetDefault();
        }

        protected virtual void SetDefault()
        {
            _xscale = new Mtblib.Graph.Component.Scale.CateScale(Mtblib.Graph.Component.ScaleDirection.X_Axis);
            _xscale.Label.Visible = false;
            _xscale.Ticks.Angle = 0;
            _title = new Mtblib.Graph.Component.Title();

            Division = 0.6;
            _boxplot.CMean.Visible = true;
            _boxplot.CMean.Size = 1.5;
            _boxplot.Mean.Visible = true;
            _boxplot.Mean.Type = 6; //solid circle
            _boxplot.Mean.Color = 64; // Indigo
            _boxplot.Mean.Size = 1.5;
            _boxplot.Individual.Visible = true;
            _boxplot.Individual.Color = 20; // Medium Gray

            _boxplotmeansymbolsize = 1.5;
            _boxplotindivsymbolsize = 1;

            _boxplot.IQRBox.Visible = false;
            _boxplot.Whisker.Visible = false;
            _boxplot.RBox.Visible = false;
            _boxplot.Outlier.Visible = false;
            _boxplot.Title.Visible = false;
            _boxplot.FigureRegion.AutoSize = false;
            _boxplot.FigureRegion.SetCoordinate(0, 1, Division, 1);
            _boxplot.DataRegion.AutoSize = false;
            _boxplot.DataRegion.SetCoordinate(0.074, 0.9641, 0.005, 0.9627); 


            _boxplot.YScale.GetCommand = () =>
            {
                #region Override GetCommand of YScale
                StringBuilder cmnd = new StringBuilder();
                if (_boxplot.YScale.LDisplay != null)
                    cmnd.AppendLine(string.Format("LDisplay {0};", string.Join(" ", _boxplot.YScale.LDisplay)));
                if (_boxplot.YScale.HDisplay != null)
                    cmnd.AppendLine(string.Format("HDisplay {0};", string.Join(" ", _boxplot.YScale.HDisplay)));
                if (_boxplot.YScale.Min < Mtblib.Tools.MtbTools.MISSINGVALUE) cmnd.AppendLine(string.Format("Min {0};", _boxplot.YScale.Min));
                if (_boxplot.YScale.Max < Mtblib.Tools.MtbTools.MISSINGVALUE) cmnd.AppendLine(string.Format("Max {0};", _boxplot.YScale.Max));

                cmnd.Append(_boxplot.YScale.Ticks.GetCommand());
                cmnd.Append(_boxplot.YScale.Refes.GetCommand());
                cmnd.Append(_boxplot.YScale.Label.GetCommand());
                if (cmnd.Length > 0) //如果有設定再加入
                    cmnd.Insert(0, string.Format("Scale {0};\r\n", (int)_boxplot.YScale.Direction));
                if (_boxplot.YScale.SecScale.Variable != null) cmnd.AppendLine("#Boxplot 不支援次座標變數設定 :(");
                return cmnd.ToString();
                #endregion
            };

            _chart.BarsRepresent = Mtblib.Graph.BarChart.Chart.ChartRepresent.A_FUNCTION_OF_A_VARIABLE;
            _chart.FuncType = Mtblib.Graph.BarChart.Chart.ChartFunctionType.SUM;
            _chart.Title.Visible = false;
            _chart.Legend.HideLegend = true;
            _chart.FigureRegion.AutoSize = false;
            _chart.FigureRegion.SetCoordinate(0, 1, 0, Division);
            _chart.DataRegion.AutoSize = false;
            _chart.DataRegion.SetCoordinate(0.074, 0.9641, 0.0735, 0.9627); //南茂建議的預設值
            _chart.YScale.GetCommand = () =>
            {
                #region Override GetCommand of YScale
                StringBuilder cmnd = new StringBuilder();
                if (BoxPlotVisibleOnly)
                {
                    cmnd.AppendLine(string.Format("LDisplay 0 0 0 0;"));
                    cmnd.AppendLine(string.Format("HDisplay 0 0 0 0;"));
                }
                else
                {
                    if (_chart.YScale.LDisplay != null)
                        cmnd.AppendLine(string.Format("LDisplay {0};", string.Join(" ", _chart.YScale.LDisplay)));
                    if (_chart.YScale.HDisplay != null)
                        cmnd.AppendLine(string.Format("HDisplay {0};", string.Join(" ", _chart.YScale.HDisplay)));
                }
                
                if (_chart.YScale.Min < Mtblib.Tools.MtbTools.MISSINGVALUE) cmnd.AppendLine(string.Format("Min {0};", _chart.YScale.Min));
                if (_chart.YScale.Max < Mtblib.Tools.MtbTools.MISSINGVALUE) cmnd.AppendLine(string.Format("Max {0};", _chart.YScale.Max));

                cmnd.Append(_chart.YScale.Ticks.GetCommand());
                cmnd.Append(_chart.YScale.Refes.GetCommand());
                cmnd.Append(_chart.YScale.Label.GetCommand());
                if (cmnd.Length > 0) //如果有設定再加入
                    cmnd.Insert(0, string.Format("Scale {0};\r\n", (int)_chart.YScale.Direction));
                if (_chart.YScale.SecScale.Variable != null) cmnd.AppendLine("#Barchart 不支援次座標變數設定 :(");
                return cmnd.ToString();
                #endregion
            };
            _chart.GraphRegion.AutoSize = false;
            _chart.GraphRegion.SetCoordinate(10, 4);

        }

        public virtual void Run()
        {
            StringBuilder cmnd = new StringBuilder();
            string macroPath = Mtblib.Tools.MtbTools.BuildTemporaryMacro("mychart.mac", GetCommand());
            cmnd.AppendLine("notitle");
            cmnd.AppendLine("brief 0");
            cmnd.AppendFormat("%\"{0}\" {1} {2};\r\n", macroPath,
                string.Join(" &\r\n", ((Mtb.Column[])VariablesAtBarChart).Select(x => x.SynthesizedName).ToArray()),
                string.Join(" &\r\n", ((Mtb.Column[])VariablesAtBoxPlot).Select(x => x.SynthesizedName).ToArray()));
            if (GroupingBy != null)
            {
                cmnd.AppendFormat("group {0};\r\n", string.Join(" &\r\n",
                    ((Mtb.Column[])GroupingBy).Select(x => x.SynthesizedName).ToArray()));
            }
            if (PanelBy != null)
            {
                cmnd.AppendFormat("pane {0};\r\n", ((Mtb.Column[])PanelBy)[0].SynthesizedName);
            }
            //string[] storeCol = Mtblib.Tools.MtbTools.CreateVariableStrArray(_ws, 1, Mtblib.Tools.MtbVarType.Column);
            //cmnd.AppendFormat("Store {0}\r\n", storeCol);
            cmnd.AppendLine(".");
            cmnd.AppendLine("title");
            cmnd.AppendLine("brief 2");
            string fpath = Mtblib.Tools.MtbTools.BuildTemporaryMacro("mycode.mtb", cmnd.ToString());

            _proj.ExecuteCommand(string.Format("exec \"{0}\" 1", fpath));

        }

        protected virtual string GetCommand()
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

            Mtb.Column[] gps = GroupingBy;
            Mtb.Column[] pane = PanelBy;

            StringBuilder cmnd = new StringBuilder();
            cmnd.AppendLine("macro");
            cmnd.AppendLine("hlbarline y trnd;");
            cmnd.AppendLine("group x.1-x.m;"); // 1~m 代表分群由外至內
            cmnd.AppendLine("pane p.");
            cmnd.AppendLine("mcolumn y trnd x.1-x.m");
            cmnd.AppendLine("mcolumn dlabmin dlabmax d.1-d.m");
            cmnd.AppendLine("mcolumn dlab dlabtrnd");
            cmnd.AppendLine("mcolumn p");

            cmnd.AppendLine("mreset");
            cmnd.AppendLine("notitle");
            cmnd.AppendLine("noecho");
            cmnd.AppendLine("brief 0");

            List<double> vlineInBarChart = new List<double>();

            #region 使用 Minitab command line 取得 BarChart 最外層的垂直線 vline(測試後比 DataTable 法快一點)
            //在262筆資料，3個分群下大概比 DataTable 法快約0.3秒
            if (gps != null && gps.Length > 1)
            {
                Console.WriteLine("{0:mm:ss:fff}\tStart vline calculation", DateTime.Now);
                StringBuilder tmpCmnd = new StringBuilder();
                int currentWsCount = _ws.Columns.Count;
                int currentConstCount = _ws.Constants.Count;
                string[] gvalCol = Mtblib.Tools.MtbTools.CreateVariableStrArray(_ws, gps.Length, Mtblib.Tools.MtbVarType.Column);
                tmpCmnd.AppendLine("notitle");
                tmpCmnd.AppendLine("brief 0");
                tmpCmnd.AppendFormat("stat {0};\r\n", varBarchart[0].SynthesizedName);
                tmpCmnd.AppendFormat("by {0};\r\n", string.Join(" &\r\n", gps.Select(x => x.SynthesizedName)));
                tmpCmnd.AppendLine("noem;");
                tmpCmnd.AppendFormat("gval {0}.\r\n", string.Join(" &\r\n", gvalCol));
                string[] strCol = Mtblib.Tools.MtbTools.CreateVariableStrArray(_ws, 2, Mtblib.Tools.MtbVarType.Column);
                string[] strConst = Mtblib.Tools.MtbTools.CreateVariableStrArray(_ws, 1, Mtblib.Tools.MtbVarType.Constant);
                tmpCmnd.AppendFormat("Count {0} {1}\r\n", gvalCol[0], strConst[0]);
                tmpCmnd.AppendFormat("Set {0}\r\n", strCol[0]);
                tmpCmnd.AppendFormat("1:{0}\r\n", strConst);
                tmpCmnd.AppendLine("End");
                //如果最外層是文字格式的資料，需要另外處理
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
                Console.WriteLine("{0:mm:ss:fff}\tvline Completed", DateTime.Now);
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
            cmnd.AppendFormat("title \"{0}\";\r\n", Title.Text == null ? "Bar-Line Plot" : Title.Text);
            if (Title.FontSize > 0) cmnd.AppendFormat("psize {0};\r\n", Title.FontSize);

            cmnd.AppendLine("offset 0 -0.045801;");
            cmnd.Append(_chart.GraphRegion.GetCommand());
            //cmnd.AppendLine("graph 10 4;");
            cmnd.AppendLine(".");

            //如果要顯示自訂Datlab位數, 就要自己算
            //如果有 panel 的做法就必須忽略
            if (pane == null && DatlabOptionAtBarChart.AutoDecimal == false)
            {
                cmnd.AppendLine("stat y;");
                cmnd.AppendLine(" by x.1-x.m;");
                //cmnd.AppendLine(" noem;");//不可因為 bar chart 使用 nomiss 就使用 noempty，因為 bar chart只是隱藏那些 bar 
                cmnd.AppendFormat(" {0} dlab.\r\n", _chart.FuncType.ToString() == "SUM" ? "SUMS" : _chart.FuncType.ToString());
                cmnd.AppendLine("fnum dlab;");
                cmnd.AppendFormat(" fixed {0}.\r\n", DatlabOptionAtBarChart.DecimalPlace);
                _chart.DataLabel.DatlabType = Mtblib.Graph.Component.Datlab.DisplayType.Column;
                _chart.DataLabel.LabelColumn = "dlab";
            }

            #region 建立 Bar chart
                cmnd.AppendFormat("chart {0}(y) &\r\n", _chart.FuncType.ToString());
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
                    switch (BarColorType)
                    {
                        case BarColorOption.ByInnerMostGroup:
                            _chart.Bar.GroupingBy = "x.m";
                            break;
                        case BarColorOption.ByOuterMostGroup:
                            _chart.Bar.GroupingBy = "x.1";
                            break;
                        case BarColorOption.Single:
                        default:
                            break;
                    }
                    //_chart.Bar.GroupingBy = "x.m";
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
                cmnd.Append(_chart.DataRegion.GetCommand());
                cmnd.Append(_chart.Title.GetCommand());
                cmnd.AppendLine("nomiss;");
                cmnd.AppendLine("noem;");
                cmnd.AppendLine("coffset 0.1;");
                cmnd.Append(_chart.GetAnnotationCommand());
                cmnd.AppendLine(".");
                #endregion
                

                if (pane == null && DatlabOptionAtBoxPlot.AutoDecimal == false)
                {
                    cmnd.AppendLine("stat trnd;");
                    cmnd.AppendLine(" by x.1-x.m;");
                    //cmnd.AppendLine(" noem;");//不可因為 bar chart 使用 nomiss 就使用 noempty，因為 bar chart只是隱藏那些 bar 
                    cmnd.AppendLine(" mean dlab.");
                    cmnd.AppendLine("fnum dlab;");
                    cmnd.AppendFormat(" fixed {0}.\r\n", DatlabOptionAtBoxPlot.DecimalPlace);
                    _boxplot.MeanDatlab.DatlabType = Mtblib.Graph.Component.Datlab.DisplayType.Column;
                    _boxplot.MeanDatlab.LabelColumn = "dlab";
                }
            //}
            #region 處理 Boxplot Individual顯示方式,包含顯示全部或部分、調整位數
            if (pane == null && _boxplot.IndivDatlab.DatlabType == Mtblib.Graph.Component.Datlab.DisplayType.YValue)
            {
                if (DatlabOptionAtBoxPlotIndiv.AutoDecimal == false) cmnd.AppendFormat(" let dlabtrnd = text(round(trnd,{0})) \r\n;", DatlabOptionAtBoxPlotIndiv.DecimalPlace);
                else cmnd.AppendFormat(" let dlabtrnd = text(round(trnd,{0})) \r\n", 3);
                _boxplot.IndivDatlab.DatlabType = Mtblib.Graph.Component.Datlab.DisplayType.Column;
                _boxplot.IndivDatlab.LabelColumn = "dlabtrnd";
            }
            else if (pane == null && _boxplot.IndivDatlab.DatlabType == Mtblib.Graph.Component.Datlab.DisplayType.Column)
            {
                cmnd.AppendLine("stat trnd;");
                cmnd.AppendLine(" by x.1-x.m;");
                //cmnd.AppendLine(" noem;");//不可因為 bar chart 使用 nomiss 就使用 noempty，因為 bar chart只是隱藏那些 bar 
                cmnd.AppendLine("Expand;");
                cmnd.AppendLine("gval d.1-d.m;");
                cmnd.AppendLine("Minimum dlabmin;");
                cmnd.AppendLine("Maximum dlabmax.");
                if (DatlabOptionAtBoxPlotIndiv.AutoDecimal == false)
                {
                    if (_maxVisible && _minVisible)
                        cmnd.AppendFormat("let dlabtrnd = if(trnd=dlabmin, text(round(trnd,{0})), IF(trnd=dlabmax, text(round(trnd,{0})), \"\" ))\r\n",
                            DatlabOptionAtBoxPlotIndiv.DecimalPlace);
                    else if (_maxVisible)
                        cmnd.AppendFormat("let dlabtrnd = if(trnd=dlabmax, text(round(trnd,{0})), \"\" )\r\n", DatlabOptionAtBoxPlotIndiv.DecimalPlace);
                    else if (_minVisible)
                        cmnd.AppendFormat("let dlabtrnd = if(trnd=dlabmin, text(round(trnd,{0})), \"\" )\r\n", DatlabOptionAtBoxPlotIndiv.DecimalPlace);
                    else;
                }
                else
                {
                    if (_maxVisible && _minVisible)
                        cmnd.AppendFormat("let dlabtrnd = if(trnd=dlabmin, text(round(trnd,{0})), IF(trnd=dlabmax, text(round(trnd,{0})), \"\" ))\r\n", 3);
                    else if (_maxVisible)
                        cmnd.AppendFormat("let dlabtrnd = if(trnd=dlabmax, text(round(trnd,{0})), \"\" )\r\n", 3);
                    else if (_minVisible)
                        cmnd.AppendFormat("let dlabtrnd = if(trnd=dlabmin, text(round(trnd,{0})), \"\" )\r\n", 3);
                    else;
                }
                //cmnd.AppendFormat("let dlabtrnd = if(trnd=dlabmin, \"\", \"\" )\r\n");
                //cmnd.AppendFormat("let dlabtrnd = if(trnd=dlabmin, MISS(), MISS() )\r\n");
                //cmnd.AppendFormat("let dlabtrnd = if(trnd=dlabmin, *, * )\r\n");

                _boxplot.IndivDatlab.DatlabType = Mtblib.Graph.Component.Datlab.DisplayType.Column;
                _boxplot.IndivDatlab.LabelColumn = "dlabtrnd"; // no need or need missing

                #region label在Minitab裡 PCR20181025, delete after test
                //將要label的值算出並儲存
                //StringBuilder cmnd22 = new StringBuilder();
                //cmnd22.AppendLine("Name c1000 \"Minimum\" c999 \"Maximum\" c998 \"MinOnly\" c997 \"MaxOnly\"");
                //for (int i = 0; i < ((Mtb.Column[])GroupingBy).Select(x => x.SynthesizedName).Count(); i++)
                //{
                //    cmnd22.AppendFormat("Name c{0} \"ByVar{1}\" \r\n", 996 - i, i + 1);
                //}
                ////cmnd22.AppendLine("Name c29 \"ByVar1\" c30 \"ByVar2\" c31 \"Minimum\" c32 \"Maximum\" c33 \"MinOnly\" c34 \"MaxOnly\"");
                //cmnd22.AppendFormat("stat {0}; \r\n", string.Join(" &\r\n", ((Mtb.Column[])VariablesAtBoxPlot).Select(x => x.SynthesizedName).ToArray()));
                //cmnd22.AppendFormat(" by {0}; \r\n", string.Join(" ", ((Mtb.Column[])GroupingBy).Select(x => x.SynthesizedName).ToArray()));
                //cmnd22.AppendLine("Expand;");
                //cmnd22.AppendFormat("gval 'ByVar{0}'-'ByVar{1}'; \r\n", 1, ((Mtb.Column[])GroupingBy).Select(x => x.SynthesizedName).Count());
                ////cmnd22.AppendFormat("gval {0}; \r\n", "'ByVar1' 'ByVar2'");
                //cmnd22.AppendLine("Minimum 'Minimum';");
                //cmnd22.AppendLine("Maximum 'Maximum'.");
                //if (_minVisible) cmnd22.AppendFormat("let MinOnly = if( {0} = Minimum, {1}, MISS()) \r\n", VariablesAtBoxPlot[0].SynthesizedName.ToString(), VariablesAtBoxPlot[0].SynthesizedName.ToString());
                //else cmnd22.AppendFormat("let MinOnly = if( {0} = Minimum, MISS(), MISS()) \r\n", VariablesAtBoxPlot[0].SynthesizedName.ToString());
                //if (_maxVisible) cmnd22.AppendFormat("let MaxOnly = if( {0} = Maximum, {1}, MISS()) \r\n", VariablesAtBoxPlot[0].SynthesizedName.ToString(), VariablesAtBoxPlot[0].SynthesizedName.ToString());
                //else cmnd22.AppendFormat("let MaxOnly = if( {0} = Maximum, MISS(), MISS()) \r\n", VariablesAtBoxPlot[0].SynthesizedName.ToString());

                //string fpath = Mtblib.Tools.MtbTools.BuildTemporaryMacro("mycode.mtb", cmnd22.ToString());
                //_proj.ExecuteCommand(string.Format("exec \"{0}\" 1", fpath));

                //List<Mtb.Column> _dataCols = new List<Mtb.Column>();

                //dynamic maxby = "MaxOnly";
                //dynamic MaxBy = Mtblib.Tools.MtbTools.GetMatchColumns(maxby, _ws);
                //Mtb.Column[] MaxByCol = (Mtb.Column[])MaxBy;
                //_dataCols.Add(MaxByCol[0]);

                //dynamic minby = "MinOnly";
                //dynamic MinBy = Mtblib.Tools.MtbTools.GetMatchColumns(minby, _ws);
                //Mtb.Column[] MinByCol = (Mtb.Column[])MinBy;
                //_dataCols.Add(MinByCol[0]);

                //System.Data.DataTable dtt = new System.Data.DataTable();
                //dtt = Mtblib.Tools.MtbTools.GetDataTableFromMtbCols(_dataCols.ToArray());
                //dtt.Columns.Add(new DataColumn("RowNum", typeof(int)));
                //dtt.Columns[0].ColumnName = "MaxOnly";
                //dtt.Columns[1].ColumnName = "MinOnly";

                //// 調整label位置
                //for (int i = 0; i < dtt.Rows.Count; i++)
                //{
                //    dtt.Rows[i]["RowNum"] = i + 1;
                //    if (_maxVisible)
                //    {
                //        if (!(dtt.Rows[i]["MaxOnly"].ToString() == Mtblib.Tools.MtbTools.MISSINGVALUE.ToString()))
                //        {
                //            Mtblib.Graph.Component.LabelPosition alabelmax = new Mtblib.Graph.Component.LabelPosition();
                //            alabelmax.Model = 1;
                //            alabelmax.RowId = i + 1;
                //            alabelmax.Placement = new double[2] { 1, 1 };
                //            _boxplot.IndivDatlab.PositionList.Add(alabelmax);
                //        }
                //    }

                //    if (_minVisible)
                //    {
                //        if (!(dtt.Rows[i]["MinOnly"].ToString() == Mtblib.Tools.MtbTools.MISSINGVALUE.ToString()))
                //        {
                //            Mtblib.Graph.Component.LabelPosition alabelmin = new Mtblib.Graph.Component.LabelPosition();
                //            alabelmin.Model = 1;
                //            alabelmin.RowId = i + 1;
                //            alabelmin.Placement = new double[2] { 1, -1 };
                //            _boxplot.IndivDatlab.PositionList.Add(alabelmin);
                //        }
                //    }
                //}
                #endregion
                #region label 跟 group 讀近c#處理
                StringBuilder cmnd22 = new StringBuilder();
                cmnd22.AppendLine("Name c1000 \"Minimum\" c999 \"Maximum\" c998 \"MinOnly\" c997 \"MaxOnly\"");
                for (int i = 0; i < ((Mtb.Column[])GroupingBy).Select(x => x.SynthesizedName).Count(); i++)
                {
                    cmnd22.AppendFormat("Name c{0} \"ByVar{1}\" \r\n", 996 - i, i + 1);
                }
                //cmnd22.AppendLine("Name c29 \"ByVar1\" c30 \"ByVar2\" c31 \"Minimum\" c32 \"Maximum\" c33 \"MinOnly\" c34 \"MaxOnly\"");
                cmnd22.AppendFormat("stat {0}; \r\n", string.Join(" &\r\n", ((Mtb.Column[])VariablesAtBoxPlot).Select(x => x.SynthesizedName).ToArray()));
                cmnd22.AppendFormat(" by {0}; \r\n", string.Join(" ", ((Mtb.Column[])GroupingBy).Select(x => x.SynthesizedName).ToArray()));
                cmnd22.AppendLine("Expand;");
                cmnd22.AppendFormat("gval 'ByVar{0}'-'ByVar{1}'; \r\n", 1, ((Mtb.Column[])GroupingBy).Select(x => x.SynthesizedName).Count());
                //cmnd22.AppendFormat("gval {0}; \r\n", "'ByVar1' 'ByVar2'");
                cmnd22.AppendLine("Minimum 'Minimum';");
                cmnd22.AppendLine("Maximum 'Maximum'.");
                if (_minVisible) cmnd22.AppendFormat("let MinOnly = if( {0} = Minimum, {1}, MISS()) \r\n", VariablesAtBoxPlot[0].SynthesizedName.ToString(), VariablesAtBoxPlot[0].SynthesizedName.ToString());
                else cmnd22.AppendFormat("let MinOnly = if( {0} = Minimum, MISS(), MISS()) \r\n", VariablesAtBoxPlot[0].SynthesizedName.ToString());
                if (_maxVisible) cmnd22.AppendFormat("let MaxOnly = if( {0} = Maximum, {1}, MISS()) \r\n", VariablesAtBoxPlot[0].SynthesizedName.ToString(), VariablesAtBoxPlot[0].SynthesizedName.ToString());
                else cmnd22.AppendFormat("let MaxOnly = if( {0} = Maximum, MISS(), MISS()) \r\n", VariablesAtBoxPlot[0].SynthesizedName.ToString());

                string fpath = Mtblib.Tools.MtbTools.BuildTemporaryMacro("mycode.mtb", cmnd22.ToString());
                _proj.ExecuteCommand(string.Format("exec \"{0}\" 1", fpath));

                List<Mtb.Column> _dataCols = new List<Mtb.Column>();

                dynamic maxby = "MaxOnly";
                dynamic MaxBy = Mtblib.Tools.MtbTools.GetMatchColumns(maxby, _ws);
                Mtb.Column[] MaxByCol = (Mtb.Column[])MaxBy;
                _dataCols.Add(MaxByCol[0]);

                dynamic minby = "MinOnly";
                dynamic MinBy = Mtblib.Tools.MtbTools.GetMatchColumns(minby, _ws);
                Mtb.Column[] MinByCol = (Mtb.Column[])MinBy;
                _dataCols.Add(MinByCol[0]);

                Mtb.Column[] GroupByCol = (Mtb.Column[])GroupingBy;
                for (int i = 0; i < GroupByCol.Length; i++) _dataCols.Add(GroupByCol[i]);

                System.Data.DataTable dtt = new System.Data.DataTable();
                dtt = Mtblib.Tools.MtbTools.GetDataTableFromMtbCols(_dataCols.ToArray());
                dtt.Columns.Add(new DataColumn("RowNum", typeof(int)));
                dtt.Columns[0].ColumnName = "MaxOnly";
                dtt.Columns[1].ColumnName = "MinOnly";
                for (int i = 0; i < GroupByCol.Length; i++)
                    dtt.Columns[2 + i].ColumnName = string.Format("GroupBy {0}", 2 + i);

                // 調整label位置
                string groupMaxused = "";
                string groupMinused = "";
                for (int i = 0; i < dtt.Rows.Count; i++)
                {
                    string groupusing = "";
                    for (int j = 0; j < GroupByCol.Length; j++) groupusing += dtt.Rows[i][string.Format("GroupBy {0}", 2 + j)].ToString();
                    dtt.Rows[i]["RowNum"] = i + 1;
                    if (_maxVisible)
                    {
                        if (!(dtt.Rows[i]["MaxOnly"].ToString() == Mtblib.Tools.MtbTools.MISSINGVALUE.ToString()))
                        {
                            if(groupusing != groupMaxused)
                            {
                                Mtblib.Graph.Component.LabelPosition alabelmax = new Mtblib.Graph.Component.LabelPosition();
                                alabelmax.Model = 1;
                                alabelmax.RowId = i + 1;
                                //double tmpMax = 0;
                                //double.TryParse(dtt.Rows[i]["MaxOnly"].ToString(), out tmpMax);
                                //if (DatlabOptionAtBoxPlot.AutoDecimal) alabelmax.Text = Math.Round(tmpMax, 3).ToString();
                                //else alabelmax.Text = Math.Round(tmpMax, DatlabOptionAtBoxPlotIndiv.DecimalPlace).ToString();
                                //alabelmax.Text = dtt.Rows[i]["MaxOnly"].ToString();
                                alabelmax.Placement = new double[2] { 1, 1 };
                                _boxplot.IndivDatlab.PositionList.Add(alabelmax);
                                groupMaxused = groupusing;
                            }
                            else
                            {
                                Mtblib.Graph.Component.LabelPosition alabelmax = new Mtblib.Graph.Component.LabelPosition();
                                alabelmax.Model = 1;
                                alabelmax.RowId = i + 1;
                                alabelmax.Text = "";
                                //alabelmax.Placement = new double[2] { 1, 1 };
                                _boxplot.IndivDatlab.PositionList.Add(alabelmax);
                            }
                        }
                    }

                    if (_minVisible)
                    {
                        if (!(dtt.Rows[i]["MinOnly"].ToString() == Mtblib.Tools.MtbTools.MISSINGVALUE.ToString()))
                        {
                            if (groupusing != groupMinused)
                            {
                                Mtblib.Graph.Component.LabelPosition alabelmin = new Mtblib.Graph.Component.LabelPosition();
                                alabelmin.Model = 1;
                                alabelmin.RowId = i + 1;
                                //double tmpMin = 0;
                                //double.TryParse(dtt.Rows[i]["MinOnly"].ToString(), out tmpMin);
                                //if (DatlabOptionAtBoxPlot.AutoDecimal) alabelmin.Text = Math.Round(tmpMin, 3).ToString();
                                //else alabelmin.Text = Math.Round(tmpMin, DatlabOptionAtBoxPlotIndiv.DecimalPlace).ToString();
                                //alabelmin.Text = dtt.Rows[i]["MinOnly"].ToString();
                                alabelmin.Placement = new double[2] { 1, -1 };
                                _boxplot.IndivDatlab.PositionList.Add(alabelmin);
                                groupMinused = groupusing;
                            }
                            else
                            {
                                Mtblib.Graph.Component.LabelPosition alabelmin = new Mtblib.Graph.Component.LabelPosition();
                                alabelmin.Model = 1;
                                alabelmin.RowId = i + 1;
                                alabelmin.Text = "";
                                //alabelmin.Placement = new double[2] { 1, -1 };
                                _boxplot.IndivDatlab.PositionList.Add(alabelmin);
                            }
                        }
                    }
                }
                #endregion
            }
            else { }
            #endregion

            #region 建立 Boxplot
            cmnd.AppendLine("boxplot trnd &");
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

            #region DataScale for tick increament 處理程序
            //指定 Tick increament 的處理程序
            if (_boxplot.YScale.Ticks.Increament < Mtblib.Tools.MtbTools.MISSINGVALUE &&
                _boxplot.YScale.Ticks.NMajor == -1)
            {
                Mtblib.Tools.GScale boxplotScale
                = Mtblib.Tools.MtbTools.GetDataScaleInCateChart(
                _proj, _ws, varBoxplot, gps, pane, "", true);
                boxplotScale.TMinimum = 5;
                boxplotScale.TMaximum = 10;
                string tickString;
                if (YScaleAtBoxPlot.Max < Mtblib.Tools.MtbTools.MISSINGVALUE && YScaleAtBoxPlot.Min < Mtblib.Tools.MtbTools.MISSINGVALUE)
                {
                    tickString = string.Format("{0}:{1}/{2}",
                        YScaleAtBoxPlot.Min, YScaleAtBoxPlot.Max, _boxplot.YScale.Ticks.Increament);
                    double autorange = (YScaleAtBoxPlot.Max - YScaleAtBoxPlot.Min) * 0.027;
                    YScaleAtBoxPlot.Max = YScaleAtBoxPlot.Max + autorange;
                    YScaleAtBoxPlot.Min = YScaleAtBoxPlot.Min - autorange;
                }
                else if (YScaleAtBoxPlot.Max == Mtblib.Tools.MtbTools.MISSINGVALUE && YScaleAtBoxPlot.Min < Mtblib.Tools.MtbTools.MISSINGVALUE)
                {
                    tickString = string.Format("{0}:{1}/{2}",
                       YScaleAtBoxPlot.Min, boxplotScale.TMaximum, _boxplot.YScale.Ticks.Increament);
                    double autorange = (boxplotScale.TMaximum - YScaleAtBoxPlot.Min) * 0.027;
                    YScaleAtBoxPlot.Min = YScaleAtBoxPlot.Min - autorange;

                }
                else if (YScaleAtBoxPlot.Max < Mtblib.Tools.MtbTools.MISSINGVALUE && YScaleAtBoxPlot.Min == Mtblib.Tools.MtbTools.MISSINGVALUE)
                {
                    tickString = string.Format("{0}:{1}/{2}",
                        boxplotScale.TMinimum, YScaleAtBoxPlot.Max, _boxplot.YScale.Ticks.Increament);
                    double autorange = (YScaleAtBoxPlot.Max - boxplotScale.TMinimum) * 0.027;
                    YScaleAtBoxPlot.Max = YScaleAtBoxPlot.Max + autorange;
                }
                else
                {
                    tickString = string.Format("{0}:{1}/{2}",
                        boxplotScale.TMinimum, boxplotScale.TMaximum, _boxplot.YScale.Ticks.Increament);
                }
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
            cmnd.Append(_boxplot.Mean.GetCommand());
            cmnd.Append(_boxplot.CMean.GetCommand());
            cmnd.Append(_boxplot.Individual.GetCommand());
            cmnd.Append(_boxplot.IQRBox.GetCommand());
            cmnd.Append(_boxplot.RBox.GetCommand());
            cmnd.Append(_boxplot.Whisker.GetCommand());
            cmnd.Append(_boxplot.Outlier.GetCommand());

            cmnd.Append(_boxplot.MeanDatlab.GetCommand());
            cmnd.Append(_boxplot.IndivDatlab.GetCommand());

            if (pane != null)
            {
                _boxplot.Panel.PaneledBy = "p";
                _boxplot.Panel.RowColumn = new int[] { 1, pane[0].GetNumDistinctRows() };
                cmnd.Append(_boxplot.Panel.GetCommand());
            }

            cmnd.Append(_boxplot.FigureRegion.GetCommand());
            cmnd.Append(_boxplot.DataRegion.GetCommand());
            cmnd.Append(_boxplot.Title.GetCommand());
            cmnd.Append(_boxplot.GetAnnotationCommand());
            cmnd.AppendLine(".");
            #endregion

            cmnd.AppendLine("endlayout");


            cmnd.AppendLine("endmacro");


            return cmnd.ToString();

        }

        private bool _boxplotvisibleonly = false;
        public bool BoxPlotVisibleOnly
        {
            get
            {
                return _boxplotvisibleonly;
            }
            set
            {
                _boxplotvisibleonly = value;
                if(value == true)
                {
                    _chart.FigureRegion.SetCoordinate(0, 1, 0, 1);
                    _boxplot.FigureRegion.SetCoordinate(0, 1, 0, 1);
                    _chart.DataRegion.SetCoordinate(0.074, 0.9641, 0.1, 0.9627);
                    _boxplot.DataRegion.SetCoordinate(0.074, 0.9641, 0.1, 0.9627);
                    YScaleAtBarChart.AxLab.Text = "";
                }
            }
        }

        protected bool _maxVisible = true;
        protected bool _minVisible = true;

        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }
        protected virtual void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Free other state (managed objects).
                _chart.Dispose();
                _boxplot.Dispose();
            }
            // Free your own state (unmanaged objects).
            // Set large fields to null.
            VariablesAtBarChart = null;
            VariablesAtBoxPlot = null;
            GroupingBy = null;
            PanelBy = null;
            _proj = null;
            _ws = null;
            GC.Collect();

        }
        ~HLBarLinePlot()
        {
            Dispose(false);
        }


    }
}
