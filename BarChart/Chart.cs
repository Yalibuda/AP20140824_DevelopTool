using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.BarChart
{
    public class Chart : IChart, IDisposable
    {
        private Mtblib.Graph.BarChart.Chart _chart;
        private Mtb.Project _proj;
        private Mtb.Worksheet _ws;
        public Chart()
        {

        }
        public Chart(Mtb.Project proj, Mtb.Worksheet ws)
        {
            SetMtbEnvironment(proj, ws);
        }
        public void SetVariables(dynamic var)
        {
            Variables = var;
        }
        protected dynamic Variables
        {
            set { _chart.Variables = value; }
            get { return _chart.Variables; }
        }

        public void SetGroupingBy(dynamic var)
        {
            GroupingBy = var;
        }
        protected dynamic GroupingBy
        {
            get
            {
                return _chart.GroupingVariables;
            }
            set
            {
                _chart.GroupingVariables = value;
            }
        }

        protected Mtb.Column[] _pane = null;
        public void SetPanelBy(dynamic var)
        {
            PanelBy = var;
        }
        /// <summary>
        /// 設定或取得分割畫面的變數，合法的輸入為單一(string/Mtb.Column)或多個(string[]/Mtb.Column[])，
        /// 可使用連續輸入表示式(string)，如 C1-C3。使用 Get 取得 Minitab 欄位陣列
        /// </summary>
        protected virtual dynamic PanelBy
        {
            set
            {
                if (value == null)
                {
                    _pane = null;
                    return;
                }
                Mtb.Column[] _cols = Mtblib.Tools.MtbTools.GetMatchColumns(value, _ws);                
                _pane = _cols;
                _chart.Panel.PaneledBy = _pane.Select(x => x.SynthesizedName).ToArray();
            }
            get
            {
                return _pane;
            }
        }

        public ChartFunctionType FuncType
        {
            get
            {
                return (BarChart.ChartFunctionType)Enum.Parse(
                    typeof(BarChart.ChartFunctionType),
                    _chart.FuncType.ToString());
            }
            set
            {
                _chart.FuncType = (Mtblib.Graph.BarChart.Chart.ChartFunctionType)Enum.Parse(
                    typeof(Mtblib.Graph.BarChart.Chart.ChartFunctionType),
                    value.ToString());
            }
        }
        public BarChart.ChartStackType StackType
        {
            set
            {
                _chart.StackType = (Mtblib.Graph.BarChart.Chart.ChartStackType)Enum.Parse(
                    typeof(Mtblib.Graph.BarChart.Chart.ChartStackType),
                    value.ToString());
            }
            get
            {
                return (BarChart.ChartStackType)Enum.Parse(
                       typeof(BarChart.ChartStackType),
                       _chart.StackType.ToString());
            }
        }

        private int colAtGroupingLv = 4;
        public int ColumnAtGroupingLevel
        {
            get
            {
                return colAtGroupingLv;
            }
            set
            {
                if (value <= 0 || value > 4) throw new ArgumentException(
                    string.Format("ColumnAtGroupingLevel:{0} 是不合法的輸入，請輸入1-4的整數。", value));
                colAtGroupingLv = value;
            }
        }

        public Component.Scale.ICateScale XScale
        {
            get { return new Component.Scale.Adapter_CateScale(_chart.XScale); }
        }

        public Component.Scale.IContScale YScale
        {
            get { return new Component.Scale.Adapter_ContScale(_chart.YScale); }
        }

        public Component.IDatlab Datlab
        {
            get { return new Component.Adapter_DatLab(_chart.DataLabel); }
        }

        public Component.Region.ILegend Legend
        {
            get { return new Component.Region.Adapter_Legend(_chart.Legend); }
        }

        public Component.ILabel Title
        {
            get { return new Component.Adapter_Lab(_chart.Title); }
        }

        public Component.IFootnote Footnotes
        {
            get { return new Component.Adapter_Footnote(_chart.FootnoteLst); }
        }

        public string GSave
        {
            get
            {
                return _chart.GraphPath;
            }
            set
            {
                _chart.GraphPath = value;
            }
        }

        public void SetMtbEnvironment(Mtb.Project proj, Mtb.Worksheet ws)
        {
            _proj = proj;
            _ws = ws;
            SetDefault();
        }

        private void SetDefault()
        {
            _chart = new Mtblib.Graph.BarChart.Chart(_proj, _ws);
            _chart.BarsRepresent = Mtblib.Graph.BarChart.Chart.ChartRepresent.A_FUNCTION_OF_A_VARIABLE;
            _chart.FuncType = Mtblib.Graph.BarChart.Chart.ChartFunctionType.SUM;
            _chart.AdjDatlabAtStackBar = true;            
            _chart.XScale.Label.Visible = false;
            _chart.Legend.Sections.Add(new Mtblib.Graph.Component.Region.LegendSection(1)
            {
                HideColumnHeader = true
            });
        }
        private string GetCommand()
        {
            /*
             * 處理的資料是 summarized data(two-way table)
             * 為了讓圖形顯示更有彈性，所以會轉置成 stacked data
             */
            if (Variables == null) throw new Exception("建立 Bar chart 指令時，未給定 Variables");

            Mtb.Column[] vars = (Mtb.Column[])Variables;
            Mtb.Column[] gps = null;
            if (_chart.GroupingVariables != null)
            {
                gps = (Mtb.Column[])_chart.GroupingVariables;
                if (gps.Length > 3) throw new ArgumentException("分群變數不可超過3。");
            }

            Mtb.Column[] pane = null;
            if (_chart.Panel.PaneledBy != null)
            {
                pane = PanelBy;               
            }

            StringBuilder cmnd = new StringBuilder(); // local macro 內容

            cmnd.AppendLine("macro");
            cmnd.AppendLine("chart y.1-y.ny;");
            cmnd.AppendLine("group x.1-x.m;");
            cmnd.AppendLine("pane p.1-p.k;"); //如果使用者有指定 panel 
            cmnd.AppendLine("datlab dlab."); //如果使用者有自己指定 column for datlab

            cmnd.AppendLine("mcolumn y.1-y.ny");
            cmnd.AppendLine("mcolumn x.1-x.m");
            cmnd.AppendLine("mcolumn p.1-p.k");
            cmnd.AppendLine("mcolumn yy ycolnm ylab stkdlab dlab xx.1-xx.m");
            cmnd.AppendLine("mcolumn pp.1-pp.k");
            cmnd.AppendLine("mconstant nn");
            cmnd.AppendLine("mreset");
            cmnd.AppendLine("notitle");
            cmnd.AppendLine("noecho");
            cmnd.AppendLine("brief 0");

            // 先堆資料
            cmnd.AppendLine("stack y.1-y.ny yy;");
            cmnd.AppendLine("subs ycolnm;");
            cmnd.AppendLine("usen.");
            cmnd.AppendLine("vorder ycolnm;");
            cmnd.AppendLine("work.");

            if (gps != null && gps.Length > 0)
            {
                cmnd.AppendLine("stack &");
                for (int i = 0; i < vars.Length; i++)
                {
                    cmnd.AppendLine("(x.1-x.m) &");
                }
                cmnd.AppendLine("(xx.1-xx.m).");
                cmnd.AppendLine("text xx.1-xx.m xx.1-xx.m");
                cmnd.AppendLine("vorder xx.1-xx.m;");
                cmnd.AppendLine("work.");
            }
            if (pane != null && pane.Length > 0)
            {
                cmnd.AppendLine("stack &");
                for (int i = 0; i < vars.Length; i++)
                {
                    cmnd.AppendLine("(p.1-p.m) &");
                }
                cmnd.AppendLine("(pp.1-pp.m).");
                cmnd.AppendLine("text pp.1-pp.m pp.1-pp.m");
                cmnd.AppendLine("vorder pp.1-pp.m;");
                cmnd.AppendLine("work.");
            }

            List<string> gpVarInMacros = new List<string>();
            List<string> gpNameInMacros = new List<string>();

            if (gps != null)
            {
                for (int i = 0; i < gps.Length; i++)
                {
                    gpVarInMacros.Add("xx." + (i + 1));
                    gpNameInMacros.Add(gps[i].Name);
                }
            }

            if (gps == null || ColumnAtGroupingLevel > gps.Length)
            {
                gpVarInMacros.Add("ycolnm");
                gpNameInMacros.Add("Datas");
            }
            else
            {
                gpVarInMacros.Insert(ColumnAtGroupingLevel - 1, "ycolnm");
                gpNameInMacros.Insert(ColumnAtGroupingLevel - 1, "Datas");
            }

            Mtblib.Graph.Component.Datlab tmpDatlab = (Mtblib.Graph.Component.Datlab)_chart.DataLabel.Clone();
            if (StackType == ChartStackType.Stack && _chart.DataLabel.Visible && _chart.AdjDatlabAtStackBar)
            {
                #region 建立 Adjust stack bar chart 要的 datlab
                cmnd.AppendLine("stat yy;");
                cmnd.AppendFormat("by {0};\r\n",string.Join(" &\r\n",gpVarInMacros));
                cmnd.AppendFormat("{0} stkdlab.\r\n",
                    FuncType.ToString().ToLower() == "sum" ? "sums" : FuncType.ToString().ToLower());
                cmnd.AppendLine("text stkdlab stkdlab");
                #endregion
                tmpDatlab.DatlabType = Mtblib.Graph.Component.Datlab.DisplayType.Column;
                tmpDatlab.LabelColumn = "stkdlab";
                tmpDatlab.Placement = new double[] { 0, -1 };
            }            

            cmnd.AppendFormat("Chart {0}(yy)*{1};\r\n",
                        FuncType.ToString(), gpVarInMacros[0]);
            if (gpVarInMacros.Count > 1)
                cmnd.AppendFormat("group {0};\r\n",
                    string.Join(" &\r\n", gpVarInMacros.Select((x, i) => new { Name = x, Index = i })
                    .Where(x => x.Index > 0).Select(x => x.Name)
                    .ToArray()));
            if (StackType == ChartStackType.Stack) cmnd.AppendLine("stack;");

            _chart.XScale.Label.MultiLables = gpNameInMacros.ToArray();
            cmnd.Append(_chart.XScale.GetCommand());


            cmnd.Append(_chart.YScale.GetCommand());
            cmnd.Append(tmpDatlab.GetCommand());
            /*
             * 對每一個 DataView 建立 Command
             * 這些處理是為了將 GroupingBy 屬性由原欄位換成 macro coded name
             */
            Mtblib.Graph.Component.DataView.DataView tmpDataview
                = (Mtblib.Graph.Component.DataView.DataView)_chart.Bar.Clone();
            tmpDataview.GroupingBy = gpVarInMacros[gpVarInMacros.Count - 1];
            cmnd.Append(tmpDataview.GetCommand());
            Mtblib.Graph.Component.MultiGraph.MPanel tmpPane = (Mtblib.Graph.Component.MultiGraph.MPanel)_chart.Panel.Clone();
            if (_chart.Panel.PaneledBy != null)
            {
                tmpPane.PaneledBy = "pp.1-pp.k";
            }

            cmnd.Append(tmpPane.GetCommand());
            cmnd.Append(_chart.Legend.GetCommand());
            cmnd.Append(_chart.GetAnnotationCommand());
            cmnd.Append(_chart.GetOptionCommand());
            cmnd.Append(_chart.GetRegionCommand());
            cmnd.AppendLine(".");
            cmnd.AppendLine("endmacro");
            return cmnd.ToString();
        }
        public void Run()
        {
            StringBuilder cmnd = new StringBuilder();
            string macroPath = Mtblib.Tools.MtbTools.BuildTemporaryMacro("mychart.mac", GetCommand());
            cmnd.AppendLine("notitle");
            cmnd.AppendLine("brief 0");
            cmnd.AppendFormat("%\"{0}\" {1};\r\n", macroPath,
                string.Join(" &\r\n", ((Mtb.Column[])Variables).Select(x => x.SynthesizedName).ToArray()));
            if (GroupingBy != null)
            {
                cmnd.AppendFormat("group {0};\r\n", string.Join(" &\r\n",
                    ((Mtb.Column[])GroupingBy).Select(x => x.SynthesizedName).ToArray()));
            }
            if (PanelBy != null)
            {
                cmnd.AppendFormat("pane {0};\r\n", ((Mtb.Column[])PanelBy).Select(x=>x.SynthesizedName).ToArray() );
            }
            cmnd.AppendLine(".");
            cmnd.AppendLine("title");
            cmnd.AppendLine("brief 2");
            string fpath = Mtblib.Tools.MtbTools.BuildTemporaryMacro("mycode.mtb", cmnd.ToString());

            _proj.ExecuteCommand(string.Format("exec \"{0}\" 1", fpath));

        }



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
            }
            // Free your own state (unmanaged objects).
            // Set large fields to null.
            Variables = null;
            GroupingBy = null;
            PanelBy = null;
            _proj = null;
            _ws = null;
            GC.Collect();

        }
        ~Chart()
        {
            Dispose(false);
        }
    }
}
