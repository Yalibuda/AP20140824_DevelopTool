using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Linq.Dynamic;
using MtbGraph.Tool;

namespace MtbGraph.TrendChart
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class GroupingTrendChart : TSPlot, IGroupingTrendChart
    {

        public GroupingTrendChart()
            : base()
        {

        }
        public GroupingTrendChart(Mtb.Project proj, Mtb.Worksheet ws)
            : base(proj, ws)
        {
            SetDefault();
        }

        protected override void SetDefault()
        {
            base.SetDefault();
            _tsplot.XScale.Label.Visible = false;
            SymbolSize = "0.75";
            _isGridXVisible = true;
            IsGridYVisible = false;
            
            #region PCR 201810XX(未開放)
            _y1Decimal = 5;
            _y2Decimal = 5;
            //_y1LabelVisible = true;
            //_y2LabelVisible = true;
            #endregion
            Datlab.Visible = Y1LabelVisible = Y2LabelVisible = true;

            #region PCR20190731
            Y1OOSDatlabColor = 8;
            Y2OOSDatlabColor = 8;
            Y1OOSSymbolColor = 8;
            Y2OOSSymbolColor = 8;
            Y1DatlabColor = COLOR[0];
            Y2DatlabColor = COLOR[1];
            #endregion
        }

        /// <summary>
        /// Avoid the default red(coding in 8)
        /// </summary>
        private int[] COLOR = { 64, 9, 12, 18, 34 }; // default 64 8 9 12 18 34
        /// <summary>
        /// All symbols are cirlce
        /// </summary>
        private int[] TYPE = { 6, 6, 6, 6, 6 };

        #region PCR 201810XX(未開放)
        /// <summary>
        /// 設定symbol Color
        /// </summary>
        public void SetSymbolColor(dynamic var)
        {
            if (var != null) _symbolcolor = var;
        }
        private int? _symbolcolor = 0;

        /// <summary>
        /// 是否只針對最後一點調整
        /// </summary>
        private bool _ifonlylastlabel = false;
        public bool IfOnlyLastLabel
        {
            set { _ifonlylastlabel = value; }
            get { return _ifonlylastlabel; }
        }

        /// <summary>
        /// All line type, change default setting
        /// </summary>
        private int[] LTYPE = { 1, 1, 1, 1, 1 };

        /// <summary>
        /// 設定Y1線型態
        /// </summary>
        /// <param name="var"></param>
        public void SetY1LineType(dynamic var)
        {
            LTYPE[0] = var;
        }
        /// <summary>
        /// 設定Y2線型態
        /// </summary>
        /// <param name="var"></param>
        public void SetY2LineType(dynamic var)
        {
            LTYPE[1] = var;
        }

        /// <summary>
        /// 設定Y1 TARGET
        /// </summary>
        private Mtb.Column[] _y1Target = null;
        private dynamic Y1Target
        {
            set
            {
                if (value == null)
                {
                    _y1Target = null;
                }
                else
                {
                    _y1Target = Mtblib.Tools.MtbTools.GetMatchColumns(value, _ws);
                }
            }
            get
            {
                return _y1Target;
            }
        }
        public void SetY1Target(dynamic var)
        {
            Y1Target = var;
        }

        ///// <summary>
        ///// 設定Y2 TARGET
        ///// </summary>
        private Mtb.Column[] _y2Target = null;
        private dynamic Y2Target
        {
            set
            {
                if (value == null)
                {
                    _y2Target = null;
                }
                else
                {
                    _y2Target = Mtblib.Tools.MtbTools.GetMatchColumns(value, _ws);
                }
            }
            get
            {
                return _y2Target;
            }
        }
        public void SetY2Target(dynamic var)
        {
            Y2Target = var;
        }
        #endregion 

        //vline is visible
        private bool IsGridXVisible
        {
            get
            {
                return _isGridXVisible;
            }
            set
            {
                _isGridXVisible = value;
            }
        }
        private bool _isGridXVisible;
        public void SetGridXVisibile(dynamic var)
        {
            _isGridXVisible = var;
        }

        //hline is visible
        private bool IsGridYVisible
        {
            get
            {
                return _isGridYVisible;
            }
            set
            {
                _isGridYVisible = value;
            }
        }
        private bool _isGridYVisible;
        public void SetGridYVisibile(dynamic var)
        {
            IsGridYVisible = var;
        }

        #region PCR 201810XX(未開放)
        /// <summary>
        /// 設定 y1 equal zero visible
        /// </summary>
        public void SetY1ZeroVisible(dynamic var)
        {
            if (var != null) _y1zerovisible = var;
        }
        private bool? _y1zerovisible = true;

        /// <summary>
        /// 設定 y1 equal zero visible
        /// </summary>
        public void SetY2ZeroVisible(dynamic var)
        {
            if (var != null) _y2zerovisible = var;
        }
        private bool? _y2zerovisible = true;

        /// <summary>
        /// OOS Symbol Size
        /// </summary>
        private string _oossymbolsize = "1.0";
        private string OOSSymbolSize
        {
            get { return _oossymbolsize; }
            set { _oossymbolsize = value; }
        }
        public void SetOOSSymbolSize(dynamic var)
        {
            OOSSymbolSize = var;
        }

        private bool nooosdatlabvisible = true;
        private bool NoOOSDatlabVisible
        {
            get { return nooosdatlabvisible; }
            set { nooosdatlabvisible = value; }
        }
        /// <summary>
        /// 設定無OOS datlab是否顯示
        /// </summary>
        public void SetNoOOSDatlabVisible(dynamic var)
        {
            NoOOSDatlabVisible = var;
        }
        ///// <summary>
        ///// OOS Symbol Size
        ///// </summary>
        //private int? _oossymbolcolor = 8;
        //private dynamic OOSSymbolColor
        //{
        //    get { return _oossymbolcolor; }
        //    set { _oossymbolcolor = value; }
        //}   
        //public void SetOOSSymbolColor(dynamic var)
        //{
        //    OOSSymbolColor = var;
        //}



        /// <summary>
        /// Y1 OOS Symbol Color
        /// </summary>
        private int? _y1oossymbolcolor;
        private dynamic Y1OOSSymbolColor
        {
            get { return _y1oossymbolcolor; }
            set { _y1oossymbolcolor = value; }
        }
        public void SetY1OOSSymbolColor(dynamic var)
        {
            Y1OOSSymbolColor = var;
        }
        /// <summary>
        /// Y2 OOS Symbol Color
        /// </summary>
        private int? _y2oossymbolcolor;
        private dynamic Y2OOSSymbolColor
        {
            get { return _y2oossymbolcolor; }
            set { _y2oossymbolcolor = value; }
        }
        public void SetY2OOSSymbolColor(dynamic var)
        {
            Y2OOSSymbolColor = var;
        }

        /// <summary>
        /// Y1 OOS Symbol Size
        /// </summary>
        private int? _y1oosdatlabcolor;
        private dynamic Y1OOSDatlabColor
        {
            get { return _y1oosdatlabcolor; }
            set { _y1oosdatlabcolor = value; }
        }
        public void SetY1OOSDatlabColor(dynamic var)
        {
            Y1OOSDatlabColor = var;
        }
        /// <summary>
        /// Y2 OOS Datlab Color
        /// </summary>
        private int? _y2oosdatlabcolor;
        private dynamic Y2OOSDatlabColor
        {
            get { return _y2oosdatlabcolor; }
            set { _y2oosdatlabcolor = value; }
        }
        public void SetY2OOSDatlabColor(dynamic var)
        {
            Y2OOSDatlabColor = var;
        }

        /// <summary>
        /// Y1 OOS Symbol Size
        /// </summary>
        private int? _y1datlabcolor;
        private dynamic Y1DatlabColor
        {
            get { return _y1datlabcolor; }
            set { _y1datlabcolor = value; }
        }
        public void SetY1DatlabColor(dynamic var)
        {
            Y1DatlabColor = var;
        }
        /// <summary>
        /// Y2 OOS Datlab Color
        /// </summary>
        private int? _y2datlabcolor;
        private dynamic Y2DatlabColor
        {
            get { return _y2datlabcolor; }
            set { _y2datlabcolor = value; }
        }
        public void SetY2DatlabColor(dynamic var)
        {
            Y2DatlabColor = var;
        }

        /// <summary>
        /// 設定Y1 COLOR
        /// </summary>
        public void SetY1Color(dynamic var)
        {
            COLOR[0] = var;
            if (!(Y1DatlabColor != 64)) Y1DatlabColor = COLOR[0];
        }
        ///<summary>
        /// 設定Y2 COLOR
        /// </summary>
        public void SetY2Color(dynamic var)
        {
            COLOR[1] = var;
            if (!(Y1DatlabColor != 9)) Y1DatlabColor = COLOR[1];
        }

        /// <summary>
        /// 設定Y1,Y2 label小數位數
        /// </summary>
        /// <param name=""></param>
        private int _y1Decimal;
        private int Y1Decimal
        {
            get
            {
                return _y1Decimal;
            }
            set
            {
                _y1Decimal = value;
                //if(int.TryParse(value.ToString(), out _y1Decimal)) 
            }
        }
        public void SetY1LabelDec(dynamic var)
        {
            _y1Decimal = var;
        }

        private int _y2Decimal;
        private int Y2Decimal
        {
            get
            {
                return _y2Decimal;
            }
            set
            {
                _y2Decimal = value;
                //if(int.TryParse(value.ToString(), out _y1Decimal)) 
            }
        }
        public void SetY2LabelDec(dynamic var)
        {
            _y2Decimal = var;
        }

        /// <summary>
        /// 設定Y1,Y2 label是否顯示
        /// </summary>
        /// <param name=""></param>
        private bool _y1LabelVisible;
        private bool Y1LabelVisible
        {
            get
            {
                return _y1LabelVisible;
            }
            set
            {
                _y1LabelVisible = value;
            }
        }
        public void SetY1LabelVisible(dynamic var)
        {
            Y1LabelVisible = var;
        }

        //Set Y2 Label Visible
        private bool _y2LabelVisible;
        private bool Y2LabelVisible
        {
            get
            {
                return _y2LabelVisible;
            }
            set
            {
                _y2LabelVisible = value;
            }
        }
        public void SetY2LabelVisible(dynamic var)
        {
            Y2LabelVisible = var;
        }
        #endregion  

        public override void SetGroupingBy(dynamic var)
        {
            //最多只能放入長度是2的變數            
            Mtb.Column[] cols = Mtblib.Tools.MtbTools.GetMatchColumns(var, _ws);
            if (cols.Length > 1)
            {
                throw new Exception("Too many columns in Group, allow at most 1 column.");
            }
            GroupingBy = var;
        }

        public void SetXGroup(dynamic var)
        {
            //最多只能放入長度是2的變數            
            Mtb.Column[] cols = Mtblib.Tools.MtbTools.GetMatchColumns(var, _ws);
            if (cols.Length > 2)
            {
                throw new Exception("Too many columns in XGroup, allow at most two columns.");
            }
            _xgroups = Mtblib.Tools.MtbTools.GetMatchColumns(var, _ws);
        }
        private Mtb.Column[] _xgroups = null;

        public void SetSymbolSize(dynamic var)
        {
            SymbolSize = var.ToString();
        }
        private string _symbolSize;
        private string SymbolSize
        {
            get { return _symbolSize; }
            set
            {
                if (value != null) _symbolSize = value;
            }
        }

        public void SetLineSize(dynamic var)
        {
            LineSize = var.ToString();
        }

        private string _lineSize;
        private string LineSize
        {
            get { return _lineSize; }
            set
            {
                if (value != null) _lineSize = value;
            }
        }

        protected override string GetCommand()
        {

            if (Variables == null) throw new Exception("建立 GroupingTrend 指令時，未給定 Variables");

            Mtb.Column[] vars = (Mtb.Column[])Variables;
            Mtb.Column[] gps = null; //用於分群的所有變數
            Mtb.Column[] gps_color = null; //要分群且分色的 groupby, 最多一個
            Mtb.Column[] gps_x = null; //要分群但不分色的 groupby,也會顯示在 stamp, 最多二個
            Mtb.Column[] stmp = null;

            if (_xgroups != null)
            {
                gps_x = (Mtb.Column[])_xgroups;
                gps = (Mtb.Column[])_xgroups;

            }
            if (GroupingBy != null)
            {
                gps_color = (Mtb.Column[])_tsplot.GroupingVariables;
                if (gps == null)
                {
                    gps = (Mtb.Column[])gps_color;
                }
                else
                {
                    List<Mtb.Column> tmpCol = gps.ToList();
                    tmpCol.AddRange((Mtb.Column[])gps_color);
                    gps = tmpCol.Distinct().ToArray();
                }
            }
            if (Stamp != null)
            {
                stmp = (Mtb.Column[])_tsplot.Stamp;
            }

            StringBuilder execcmd = new StringBuilder();
            execcmd.AppendLine("Sort;");
            execcmd.AppendFormat("  By ");
            for (int i = 0; i < _xgroups.Count(); i++) execcmd.AppendFormat("'{0}' ", _xgroups[i].Name);
            execcmd.AppendLine(";");
            execcmd.AppendLine("  Original.");
            string fpathtmp = Mtblib.Tools.MtbTools.BuildTemporaryMacro("exectmp.mtb", execcmd.ToString());
            _proj.ExecuteCommand(string.Format("exec \"{0}\" 1", fpathtmp));

            StringBuilder cmnd = new StringBuilder();

            cmnd.AppendLine("macro");
            cmnd.AppendLine("mytrend y.1-y.ny;");
            cmnd.AppendLine("groupby g;");
            cmnd.AppendLine("xgroup x.1-x.m;");
            cmnd.AppendLine("stamp stmp");

            cmnd.AppendLine("mcolumn y.1-y.ny ");
            cmnd.AppendLine("mcolumn x.1-x.m g stmp");
            cmnd.AppendLine("mcolumn xx.1-xx.4 nn.1-nn.4 yy uniq freq index link vline");
            cmnd.AppendLine("mcolumn colr symb line type lcolr lsymb lline ltype tmp.1-tmp.10");
            cmnd.AppendLine("mconstant i j cnt ycnt const.1-const.5");

            cmnd.AppendLine("mreset");
            cmnd.AppendLine("notitle");
            cmnd.AppendLine("noecho");
            cmnd.AppendLine("brief 0");

            #region 定義 col 與 macro 變數字典
            //定出 variables 在 macros 的名稱 
            Dictionary<Mtb.Column, string> dcVar = new Dictionary<Mtb.Column, string>();
            for (int i = 0; i < vars.Length; i++)
            {
                Mtb.Column col = vars[i];
                dcVar.Add(col, "y." + (i + 1));
            }
            // 定出 group_color 在 macro 的名稱
            Dictionary<Mtb.Column, string> dcGroup_color = new Dictionary<Mtb.Column, string>();
            if (gps_color != null)
            {
                dcGroup_color.Add(gps_color[0], "g");
            }
            // 定出 xgroup 在 macro 的名稱
            Dictionary<Mtb.Column, string> dcGroup_x = new Dictionary<Mtb.Column, string>();
            if (gps_x != null)
            {
                for (int i = 0; i < gps_x.Length; i++)
                {
                    Mtb.Column col = gps_x[i];
                    dcGroup_x.Add(col, "x." + (i + 1));
                }
            }

            // 定出 groupingby 在 macros 的名稱, group_color&group_x 可能有重複, 所以以 group_color 優先
            // 這樣處理是避免分群時候重複用到相同的欄位
            Dictionary<Mtb.Column, string> dcGroup = new Dictionary<Mtb.Column, string>();
            if (gps != null)
            {
                string macLab;
                foreach (var col in gps)
                {
                    if (dcGroup_color.Count > 0 && dcGroup_color.TryGetValue(col, out macLab))
                    {
                        dcGroup.Add(col, macLab);
                    }
                    else if (dcGroup_x.Count > 0 && dcGroup_x.TryGetValue(col, out macLab))
                    {
                        dcGroup.Add(col, macLab);
                    }
                }
            }
            // 定出 stamp 在 macros 的名稱
            Dictionary<Mtb.Column, string> dcStamp = new Dictionary<Mtb.Column, string>();
            if (stmp != null)
            {
                dcStamp.Add(stmp[0], "stmp");
            }
            #endregion

            // 做出 dummy stamp, 根據 xgroup
            // 從最外層開始，用 Stats By, 依 By A, By A&B, By A&B&C 的方式逐一找出位置
            // 而且要留意 value ordering 在這裡要依 worksheet 順序 
            #region 設定 dummy column
            if (dcGroup_x.Count > 0)
            {
                cmnd.AppendLine("vorder x.1-x.m;");
                cmnd.AppendLine("work.");
                cmnd.AppendLine("let yCnt = count(y.1)");
                cmnd.AppendLine("set index");
                cmnd.AppendLine(" 1:yCnt");
                cmnd.AppendLine("end");

                for (int i = 0; i < dcGroup_x.Count; i++)
                {
                    cmnd.AppendLine("stat y.1;");
                    cmnd.AppendFormat(" by {0};\r\n", i == 0 ? "x.1" : string.Format("x.1-x.{0}", i + 1));
                    cmnd.AppendFormat(" gval {0};\r\n", i == 0 ? "tmp.1" : string.Format("tmp.1-tmp.{0}", i + 1));
                    cmnd.AppendLine(" Count nn.1.");

                    cmnd.AppendLine("let nn.1 = parsum(nn.1)");
                    if (i == 0) //用最外層的 XGroup 取分割線
                    {
                        cmnd.AppendLine("copy nn.1 vline");
                        cmnd.AppendLine("let vline = vline +0.5");
                        cmnd.AppendLine("let const.1 = Count(vline)");
                        cmnd.AppendLine("delete const.1 vline");
                    }
                    cmnd.AppendLine("let nn.2 = lag(nn.1)");
                    cmnd.AppendLine("let nn.2[1] = 1");
                    cmnd.AppendLine("let nn.1 = floor((nn.1+nn.2)/2)");
                    cmnd.AppendFormat("convert nn.1 tmp.{0} index xx.{0}\r\n", i + 1);                    
                }
            }
            #endregion

            // 決定 color/ symbol 的設計
            // 如果 y 的數量>1, 要先堆疊資料
            if (dcGroup.Count > 0)
            {
                string[] var_suffix_in_macro = dcVar.Values.ToArray();
                #region 處理 Y 的 label
                if (dcVar.Count > 1)
                {
                    cmnd.AppendLine("stack &");
                    string stack_group_string = string.Join(" ", dcGroup.Values.ToArray());
                    for (int i = 0; i < dcVar.Count; i++)
                    {
                        cmnd.AppendFormat("( {0} {1} ) &\r\n", var_suffix_in_macro[i], stack_group_string);
                    }
                    cmnd.AppendFormat("(yy tmp.2-tmp.{0}); \r\n", dcGroup.Count + 1);
                    cmnd.AppendLine("subs tmp.1;");
                    cmnd.AppendLine("usename.");
                }
                else
                {
                    string copy_group_string = string.Join(" ", dcGroup.Values.ToArray());
                    cmnd.AppendLine("tset tmp.1");
                    cmnd.AppendFormat("(\"{0}\"){1} \r\n", vars[0].Name, vars[0].RowCount);
                    cmnd.AppendLine("end");
                    cmnd.AppendFormat("copy y.1 {0} yy tmp.2-tmp.{1}\r\n", copy_group_string, dcGroup.Count + 1);
                }
                #endregion

                //決定 g 對應的欄位, 通常是最後一個，因為不確定使用者如何定義, 所以還是保險抓一下正確位置
                int index_of_gps_color_in_gps = -1;
                if (dcGroup_color.Count > 0 )
                {
                    index_of_gps_color_in_gps = Array.IndexOf(gps, gps_color[0]);
                }

                //針對 tmp.1-tmp.XX 開始處理
                // 在 #y=1 & #group_color=0 就單一顏色
                #region 取得 legend 顯示設定
                cmnd.AppendLine("stat yy;");
                cmnd.AppendFormat(" by tmp.1{0};\r\n",
                    index_of_gps_color_in_gps >= 0 ? string.Format(" tmp.{0}", index_of_gps_color_in_gps + 2) : "");
                cmnd.AppendFormat(" gval tmp.5{0}.\r\n",
                    index_of_gps_color_in_gps > 0 ? string.Format(" tmp.{0}", index_of_gps_color_in_gps + 6) : "");

                if (index_of_gps_color_in_gps >= 0)
                {
                    cmnd.AppendFormat("let uniq = conc(tmp.5,tmp.{0})\r\n", index_of_gps_color_in_gps + 6);
                }
                else
                {
                    cmnd.AppendLine("copy tmp.5 uniq");
                }

                cmnd.AppendLine("let cnt = count(tmp.5)"); //決定要顯色的組數
                cmnd.AppendLine("let const.1 = cnt + 1");
                //Set color list 
                cmnd.AppendLine("set colr");
                cmnd.AppendFormat(" cnt({0})\r\n", string.Join(" ", COLOR)); 
                cmnd.AppendLine("end");
                cmnd.AppendLine("let const.2 = Count(colr)");
                cmnd.AppendLine("delete const.1:const.2 colr");
                //Set symbol list 
                cmnd.AppendLine("set symb");
                cmnd.AppendFormat(" cnt({0})\r\n", string.Join(" ", Mtblib.Tools.MtbTools.SYMBOLTYPE));
                cmnd.AppendLine("end");
                cmnd.AppendLine("let const.2 = Count(symb)");
                cmnd.AppendLine("delete const.1:const.2 symb");
                //Set line list 
                cmnd.AppendLine("set line");
                cmnd.AppendFormat(" cnt({0})\r\n", string.Join(" ", LTYPE)); //Mtblib.Tools.MtbTools.LINETPYE
                cmnd.AppendLine("end");
                cmnd.AppendLine("let const.2 = Count(line)");
                cmnd.AppendLine("delete const.1:const.2 line");
                //Set type list
                cmnd.AppendLine("set type");
                cmnd.AppendFormat(" cnt({0})\r\n", string.Join(" ", TYPE));
                cmnd.AppendLine("end");
                cmnd.AppendLine("let const.2 = Count(type)");
                cmnd.AppendLine("delete const.1:const.2 type");

                #endregion

                //計算出 legend box 的主體，取得實際要使用的顯示風格
                #region 取得實際要使用的顯示
                cmnd.AppendLine("stat yy;");
                cmnd.AppendFormat(" by tmp.1-tmp.{0};\r\n", dcGroup.Count + 1);
                cmnd.AppendFormat(" gval tmp.5-tmp.{0}.\r\n", dcGroup.Count + 5);

                cmnd.AppendFormat("let link = {0}\r\n",
                    dcGroup_color.Count > 0 ? string.Format("conc(tmp.5, tmp.{0})", index_of_gps_color_in_gps + 6) : "tmp.5");  //訂出對應 colr, symb, line 的 key
                cmnd.AppendLine("convert uniq colr link lcolr");
                cmnd.AppendLine("convert uniq symb link lsymb");
                cmnd.AppendLine("convert uniq line link lline");
                cmnd.AppendLine("convert uniq type link ltype");
                #endregion
            }
            //開始畫圖
            cmnd.AppendLine("tsplot y.1-y.ny;");
            cmnd.AppendLine("overl;");
            if (dcGroup.Count > 0 || dcStamp.Count > 0)
            {
                string stmp_cmnd = "stamp ";
                if (dcStamp.Count > 0) stmp_cmnd = stmp_cmnd + "stmp &\r\n";

                if (dcGroup_x.Count > 0)
                {
                    stmp_cmnd = stmp_cmnd + string.Format("xx.{0}-xx.1", dcGroup_x.Count);
                }
                stmp_cmnd = stmp_cmnd + ";";
                cmnd.AppendLine(stmp_cmnd);

            }
            _tsplot.XScale.Ticks.SetTicks(string.Format("1:{0}", vars[0].RowCount));
            if (_tsplot.XScale.Label.Visible)
            {
                _tsplot.XScale.Label.MultiLables = gps_x.Select(x => x.Name).ToArray();
            }
            if (dcGroup_x.Count > 0 && IsGridXVisible)
            {
                _tsplot.XScale.Refes.Values = "vline";
                _tsplot.XScale.Refes.Color = new string[] { "20" };
            }
            cmnd.Append(_tsplot.XScale.GetCommand());

            if (IsGridYVisible) cmnd.AppendLine("Grid 2;");
            Mtblib.Graph.Component.Scale.ContScale tmpYScale = (Mtblib.Graph.Component.Scale.ContScale)_tsplot.YScale.Clone();
            if (_tsplot.YScale.SecScale.Variable != null)
            {
                string[] sec_scale_var = _tsplot.YScale.SecScale.Variable;
                string[] sec_scale_var_in_macro = new string[sec_scale_var.Length];
                for (int i = 0; i < sec_scale_var.Length; i++)
                {
                    dcVar.TryGetValue(_ws.Columns.Item(sec_scale_var[i]), out sec_scale_var_in_macro[i]);
                }
                tmpYScale.SecScale.Variable = sec_scale_var_in_macro;
            }
            cmnd.Append(tmpYScale.GetCommand());


            #region PCR 201810XX(待測試)
            // 處理 y label相關，包含Y1, Y2是否顯示、調整小數位數、超過Target顯示，尚未開放 PCR 201810XX
            DataTable dttmp = Mtblib.Tools.MtbTools.GetDataTableFromMtbCols(vars);
            DataTable dty1target = null;
            DataTable dty2target = null;
            if (_y1Target != null) dty1target = Mtblib.Tools.MtbTools.GetDataTableFromMtbCols(_y1Target); //更換取資料方式
            //Mtb.Column[] tmpCol1 = (Mtb.Column[])Y1Target;
            //dty1target = Mtblib.Tools.MtbTools.GetDataTableFromMtbCols(tmpCol1);

            if (_y2Target != null) dty2target = Mtblib.Tools.MtbTools.GetDataTableFromMtbCols(_y2Target);

            #region get grouping id and row id

            execcmd.Clear();
            execcmd.AppendLine("Erase C990-C1000.");
            execcmd.AppendLine("Name C1000 'RowId_tmp'");
            execcmd.AppendLine("Set 'RowId_tmp'");
            execcmd.AppendFormat("1(1 : {0} / 1)1 \r\n", dttmp.Rows.Count);
            execcmd.AppendLine("End.");
            for (int i = 0; i < _xgroups.Count(); i++)
                execcmd.AppendFormat("Name C{0} '{1}_tmp'\r\n", 999 - i, _xgroups[i].Name);
            execcmd.AppendFormat("Name C{0} 'GroupingId_tmp'\r\n", 999 - _xgroups.Count());
            execcmd.AppendFormat("Statistics '{0}';\r\n", vars[0].Name);
            execcmd.AppendFormat("By ");
            for (int i = 0; i < _xgroups.Count(); i++) execcmd.AppendFormat("'{0}' ", _xgroups[i].Name);
            execcmd.Append(";\r\n");
            execcmd.AppendLine("Expand;");
            execcmd.AppendFormat("GValues ");
            for (int i = 0; i < _xgroups.Count(); i++) execcmd.AppendFormat("'{0}_tmp' ", _xgroups[i].Name);
            execcmd.Append(";\r\n");
            execcmd.AppendLine("CumN 'GroupingId_tmp'.");

            fpathtmp = Mtblib.Tools.MtbTools.BuildTemporaryMacro("exectmp.mtb", execcmd.ToString());
            _proj.ExecuteCommand(string.Format("exec \"{0}\" 1", fpathtmp));

            Mtb.Column[] rowid_column = Mtblib.Tools.MtbTools.GetMatchColumns("C1000", _ws);
            DataTable rowid = Mtblib.Tools.MtbTools.GetDataTableFromMtbCols(rowid_column);
            string groupingidcolumn = string.Format("C{0}", 999 - _xgroups.Count());
            Mtb.Column[] groupingid_column = Mtblib.Tools.MtbTools.GetMatchColumns(groupingidcolumn, _ws);
            DataTable groupingid = Mtblib.Tools.MtbTools.GetDataTableFromMtbCols(groupingid_column);

            //ERASE 'GroupingId_tmp'

            #endregion

            List<Mtblib.Graph.Component.DataView.DataViewPosition> dataViewPositionsList = new List<Mtblib.Graph.Component.DataView.DataViewPosition>();
            Mtblib.Graph.Component.DataView.DataViewPosition dataViewPositionItem = new Mtblib.Graph.Component.DataView.DataViewPosition(Mtblib.Graph.Component.DataView.DataViewPosition.DataViewType.Symbol);
            if (Datlab.Visible == false) Y1LabelVisible = Y2LabelVisible = false;
            float rawSymbolSize;
            float.TryParse(SymbolSize, out rawSymbolSize);
            float oosSymbolSize;
            float.TryParse(OOSSymbolSize, out oosSymbolSize);
            //add datlab.visible = true execute
            for (int i = 0; i < dttmp.Rows.Count; i++)
            {
                double decimaltmp = 1;
                Mtblib.Graph.Component.LabelPosition labelpositionitem = new Mtblib.Graph.Component.LabelPosition();
                string labelstring = "";
                labelpositionitem.Model = 1;
                labelpositionitem.RowId = i + 1;
                if (Y1LabelVisible == true) // now follow datlab.visible
                {
                    
                    labelpositionitem.FontColor = Y1DatlabColor; // default 64, light green 
                    for (int j = 0; j < Y1Decimal; j++) decimaltmp *= 10;
                    labelstring = (Math.Round(decimaltmp * (double)dttmp.Rows[i].ItemArray[0]) / decimaltmp).ToString();

                    if (_y1Target != null && (double)dty1target.Rows[i].ItemArray[0] != 0) //有目標且目標不為0
                    {
                        if (_ifonlylastlabel) // only last visible
                        {
                            //非分群最後一個 label ""
                            if (rowid.Rows[i].ItemArray[0].ToString() != groupingid.Rows[i].ItemArray[0].ToString()) labelstring = "";
                            else //分群最後一個label
                            {
                                labelpositionitem.FontColor =
                                    ((double)dttmp.Rows[i].ItemArray[0] > (double)dty1target.Rows[i].ItemArray[0]) ? Y1OOSDatlabColor : Y1DatlabColor;
                                if ((double)dttmp.Rows[i].ItemArray[0] == 0 & _y1zerovisible == false) labelstring = "";
                                //if oos, change color and size
                                dataViewPositionItem = new Mtblib.Graph.Component.DataView.DataViewPosition(Mtblib.Graph.Component.DataView.DataViewPosition.DataViewType.Symbol);
                                dataViewPositionItem.Model = labelpositionitem.Model;
                                dataViewPositionItem.RowId = i + 1;
                                dataViewPositionItem.Size =
                                    ((double)dttmp.Rows[i].ItemArray[0] > (double)dty1target.Rows[i].ItemArray[0]) ? oosSymbolSize : rawSymbolSize;//user define
                                if (_symbolcolor != 0)
                                {
                                    dataViewPositionItem.Color =
                                    ((double)dttmp.Rows[i].ItemArray[0] > (double)dty1target.Rows[i].ItemArray[0]) ? Y1OOSSymbolColor : _symbolcolor;
                                }
                                else
                                {
                                    dataViewPositionItem.Color =
                                    ((double)dttmp.Rows[i].ItemArray[0] > (double)dty1target.Rows[i].ItemArray[0]) ? Y1OOSSymbolColor : COLOR[0];
                                }
                                dataViewPositionsList.Add(dataViewPositionItem);
                            }
                        }
                        else
                        {
                            if ((double)dttmp.Rows[i].ItemArray[0] > (double)dty1target.Rows[i].ItemArray[0]) //oos
                            {
                                labelpositionitem.FontColor = Y1OOSDatlabColor; //改顏色? user define
                                if ((double)dttmp.Rows[i].ItemArray[0] == 0 & _y1zerovisible == false) labelstring = "";
                                    //調整color and symbol
                                dataViewPositionItem = new Mtblib.Graph.Component.DataView.DataViewPosition(Mtblib.Graph.Component.DataView.DataViewPosition.DataViewType.Symbol);
                                dataViewPositionItem.Model = labelpositionitem.Model;
                                dataViewPositionItem.RowId = i + 1;
                                dataViewPositionItem.Size =
                                    ((double)dttmp.Rows[i].ItemArray[0] > (double)dty1target.Rows[i].ItemArray[0]) ? oosSymbolSize : rawSymbolSize;//user define
                                if (_symbolcolor != 0)
                                {
                                    dataViewPositionItem.Color =
                                   ((double)dttmp.Rows[i].ItemArray[0] > (double)dty1target.Rows[i].ItemArray[0]) ? Y1OOSSymbolColor : _symbolcolor;
                                }
                                else
                                {
                                    dataViewPositionItem.Color =
                                   ((double)dttmp.Rows[i].ItemArray[0] > (double)dty1target.Rows[i].ItemArray[0]) ? Y1OOSSymbolColor : COLOR[0];
                                }
                                dataViewPositionsList.Add(dataViewPositionItem);
                            }
                            else if (!NoOOSDatlabVisible) labelstring = "";
                            else if ((double)dttmp.Rows[i].ItemArray[0] == 0 & _y1zerovisible == false) labelstring = "";
                        }
                        labelpositionitem.Text = labelstring;
                        _tsplot.DataLabel.PositionList.Add(labelpositionitem);
                        //if (((double)dty1target.Rows[i].ItemArray[0]) != 0)
                        //{
                        //    if ((double)dttest.Rows[i].ItemArray[0] > (double)dty1target.Rows[i].ItemArray[0]) //oos
                        //    {
                        //        //非分群最後一個 label ""
                        //        if (_ifonlylastlabel) if (rowid.Rows[i].ItemArray[0].ToString() != groupingid.Rows[i].ItemArray[0].ToString()) tmpstring = "";

                        //    }
                        //    else //no oos
                        //    {
                        //        tmpstring = "";
                        //    }
                        //}
                        //else //target zero, no compare
                        //{
                        //    tmpstring = ""; // no spec, unimportant, no label
                        //}
                    }
                    else //no target
                    {
                        if (_ifonlylastlabel)
                            //非分群最後一個 label ""
                            if (rowid.Rows[i].ItemArray[0].ToString() != groupingid.Rows[i].ItemArray[0].ToString()) labelstring = "";
                        if ((double)dttmp.Rows[i].ItemArray[0] == 0 & _y1zerovisible == false) labelstring = "";
                        labelpositionitem.Text = labelstring;
                        _tsplot.DataLabel.PositionList.Add(labelpositionitem);
                    }

                    #region (waiting for close)
                    //if (_ifonlylastlabel) // if only last label
                    //{
                    //    if (rowid.Rows[i].ItemArray[0].ToString() == groupingid.Rows[i].ItemArray[0].ToString())
                    //    {
                    //        //顯示 datlab
                    //        testitem.Model = 1;
                    //        testitem.Text = tmpstring;
                    //        testitem.RowId = i + 1;
                    //        testitem.FontColor = 50;
                    //        _tsplot.DataLabel.PositionList.Add(testitem);

                    //        //調整color and symbol
                    //        dataViewPositionItem = new Mtblib.Graph.Component.DataView.DataViewPosition(Mtblib.Graph.Component.DataView.DataViewPosition.DataViewType.Symbol);
                    //        dataViewPositionItem.Model = 1;
                    //        dataViewPositionItem.Color = 8; // 8 is red, default, user could change
                    //        dataViewPositionItem.RowId = i + 1;
                    //        dataViewPositionItem.Size = 1.5;//user define
                    //        dataViewPositionsList.Add(dataViewPositionItem);
                    //    }
                    //    else
                    //    {
                    //        tmpstring = "";

                    //        testitem.Model = 1;
                    //        testitem.Text = tmpstring;
                    //        testitem.RowId = i + 1;
                    //        //testitem.FontColor = 50;
                    //        _tsplot.DataLabel.PositionList.Add(testitem);
                    //    }
                    //}
                    
                    //else // oos 全部顯示 or no target 全部顯示
                    //{
                    //    //顯示 datlab
                    //    testitem.Model = 1;
                    //    testitem.Text = tmpstring;
                    //    testitem.RowId = i + 1;
                    //    testitem.FontColor = 50;
                    //    _tsplot.DataLabel.PositionList.Add(testitem);

                    //    //調整color and symbol
                    //    dataViewPositionItem = new Mtblib.Graph.Component.DataView.DataViewPosition(Mtblib.Graph.Component.DataView.DataViewPosition.DataViewType.Symbol);
                    //    dataViewPositionItem.Model = 1;
                    //    dataViewPositionItem.Color = 8; // 8 is red, default, user could change
                    //    dataViewPositionItem.RowId = i + 1;
                    //    dataViewPositionItem.Size = 1.5;//user define
                    //    dataViewPositionsList.Add(dataViewPositionItem);
                    //}
                    #endregion
                }
                else // 非多餘,再Y2顯示關閉Y1顯示時需要
                {
                    labelpositionitem.Text = labelstring;
                    _tsplot.DataLabel.PositionList.Add(labelpositionitem);
                }


                decimaltmp = 1;
                labelpositionitem = new Mtblib.Graph.Component.LabelPosition();
                labelstring = "";
                labelpositionitem.Model = 2;
                labelpositionitem.RowId = i + 1;
                //Y2LabelVisible = false; // for testing 
                if (Y2LabelVisible == true)
                {
                    labelpositionitem.FontColor = Y2DatlabColor; 
                    for (int j = 0; j < Y2Decimal; j++) decimaltmp *= 10;
                    labelstring = (Math.Round(decimaltmp * (double)dttmp.Rows[i].ItemArray[1]) / decimaltmp).ToString();

                    if (_y2Target != null && (double)dty2target.Rows[i].ItemArray[0] != 0) //有目標
                    {
                        if (_ifonlylastlabel) // only last visible
                        {
                            //非分群最後一個 label ""
                            if (rowid.Rows[i].ItemArray[0].ToString() != groupingid.Rows[i].ItemArray[0].ToString()) labelstring = "";
                            else //分群最後一個label
                            {
                                labelpositionitem.FontColor =
                                    ((double)dttmp.Rows[i].ItemArray[0] > (double)dty2target.Rows[i].ItemArray[0]) ? Y2OOSDatlabColor : Y2DatlabColor;
                                if ((double)dttmp.Rows[i].ItemArray[0] == 0 & _y2zerovisible == false) labelstring = "";
                                //if oos, change color and size
                                dataViewPositionItem = new Mtblib.Graph.Component.DataView.DataViewPosition(Mtblib.Graph.Component.DataView.DataViewPosition.DataViewType.Symbol);
                                dataViewPositionItem.Model = labelpositionitem.Model;
                                dataViewPositionItem.RowId = i + 1;
                                dataViewPositionItem.Size =
                                    ((double)dttmp.Rows[i].ItemArray[0] > (double)dty2target.Rows[i].ItemArray[0]) ? oosSymbolSize : rawSymbolSize;//user define
                                if (_symbolcolor != 0)
                                {
                                    dataViewPositionItem.Color =
                                        ((double)dttmp.Rows[i].ItemArray[0] > (double)dty2target.Rows[i].ItemArray[0]) ? Y2OOSSymbolColor : _symbolcolor;
                                }
                                else
                                {
                                    dataViewPositionItem.Color =
                                        ((double)dttmp.Rows[i].ItemArray[0] > (double)dty2target.Rows[i].ItemArray[0]) ? Y2OOSSymbolColor : COLOR[1];
                                }
                                dataViewPositionsList.Add(dataViewPositionItem);
                            }
                        }
                        else
                        {
                            if ((double)dttmp.Rows[i].ItemArray[0] > (double)dty2target.Rows[i].ItemArray[0]) //oos
                            {
                                labelpositionitem.FontColor = Y2OOSDatlabColor; //改顏色? user define
                                if ((double)dttmp.Rows[i].ItemArray[0] == 0 & _y2zerovisible == false) labelstring = "";
                                //調整color and symbol
                                dataViewPositionItem = new Mtblib.Graph.Component.DataView.DataViewPosition(Mtblib.Graph.Component.DataView.DataViewPosition.DataViewType.Symbol);
                                dataViewPositionItem.Model = labelpositionitem.Model;
                                dataViewPositionItem.RowId = i + 1;
                                dataViewPositionItem.Size =
                                    ((double)dttmp.Rows[i].ItemArray[1] > (double)dty2target.Rows[i].ItemArray[0]) ? oosSymbolSize : rawSymbolSize;//user define
                                if (_symbolcolor != 0)
                                {
                                    dataViewPositionItem.Color =
                                   ((double)dttmp.Rows[i].ItemArray[1] > (double)dty2target.Rows[i].ItemArray[0]) ? Y2OOSSymbolColor : _symbolcolor;

                                }
                                else
                                {
                                    dataViewPositionItem.Color =
                                   ((double)dttmp.Rows[i].ItemArray[1] > (double)dty2target.Rows[i].ItemArray[0]) ? Y2OOSSymbolColor : COLOR[1];
                                }
                                dataViewPositionsList.Add(dataViewPositionItem);
                            }
                            else if (!NoOOSDatlabVisible) labelstring = "";
                            else if ((double)dttmp.Rows[i].ItemArray[1] == 0 & _y2zerovisible == false) labelstring = "";
                        }
                        labelpositionitem.Text = labelstring;
                        _tsplot.DataLabel.PositionList.Add(labelpositionitem);
                    }
                    else //no target
                    {
                        if (_ifonlylastlabel)
                            //非分群最後一個 label ""
                            if (rowid.Rows[i].ItemArray[0].ToString() != groupingid.Rows[i].ItemArray[0].ToString()) labelstring = "";
                        if ((double)dttmp.Rows[i].ItemArray[0] == 0 & _y2zerovisible == false) labelstring = "";
                        labelpositionitem.Text = labelstring;
                        _tsplot.DataLabel.PositionList.Add(labelpositionitem);
                    }

                    #region raw waiting for delete
                    //for (int j = 0; j < Y2Decimal; j++) decimaltmp *= 10;
                    //tmpstring = (Math.Round(decimaltmp * (double)dttest.Rows[i].ItemArray[1]) / decimaltmp).ToString(); ;
                    //if (_y2Target != null)
                    //{
                    //    tmpstring =
                    //    ((double)dttest.Rows[i].ItemArray[1] > (double)dty2target.Rows[i].ItemArray[0]) ? tmpstring : "";
                    //}
                    //if (_ifonlylastlabel)
                    //{
                    //    if (rowid.Rows[i].ItemArray[0].ToString() == groupingid.Rows[i].ItemArray[0].ToString())
                    //    {
                    //        testitem = new Mtblib.Graph.Component.LabelPosition();
                    //        testitem.Model = 2;
                    //        testitem.Text = tmpstring;
                    //        testitem.RowId = i + 1;
                    //        testitem.FontColor = 4;
                    //        _tsplot.DataLabel.PositionList.Add(testitem);
                    //    }
                    //    else { }
                    //}
                    //else
                    //{
                    //    testitem = new Mtblib.Graph.Component.LabelPosition();
                    //    testitem.Model = 2;
                    //    testitem.Text = tmpstring;
                    //    testitem.RowId = i + 1;
                    //    testitem.FontColor = 4;
                    //    _tsplot.DataLabel.PositionList.Add(testitem);
                    //}
                    #endregion
                }
                else
                {
                    labelpositionitem.Text = labelstring;
                    _tsplot.DataLabel.PositionList.Add(labelpositionitem);
                }
            }

            
            #region raw 保留一份
            //for (int i = 0; i < dttest.Rows.Count; i++)
            //{
            //    double decimaltmp = 1;
            //    Mtblib.Graph.Component.LabelPosition testitem = new Mtblib.Graph.Component.LabelPosition();
            //    string tmpstring = "";
            //    if (Y1LabelVisible == true)
            //    {
            //        for (int j = 0; j < Y1Decimal; j++) decimaltmp *= 10;
            //        tmpstring = (Math.Round(decimaltmp * (double)dttest.Rows[i].ItemArray[0]) / decimaltmp).ToString();
            //        if (_y1Target != null)
            //        {
            //            tmpstring =
            //            ((double)dttest.Rows[i].ItemArray[0] > (double)dty1target.Rows[i].ItemArray[0]) ? tmpstring : "";
            //        }
            //        testitem.Model = 1;
            //        testitem.Text = tmpstring;
            //        testitem.RowId = i + 1;
            //        testitem.FontColor = 50;
            //        _tsplot.DataLabel.PositionList.Add(testitem);
            //    }
            //    else
            //    {
            //        tmpstring = "";

            //        testitem.Model = 1;
            //        testitem.Text = tmpstring;
            //        testitem.RowId = i + 1;
            //        _tsplot.DataLabel.PositionList.Add(testitem);
            //    }

            //    if (Y2LabelVisible == true)
            //    {
            //        decimaltmp = 1;
            //        for (int j = 0; j < Y2Decimal; j++) decimaltmp *= 10;
            //        tmpstring = (Math.Round(decimaltmp * (double)dttest.Rows[i].ItemArray[1]) / decimaltmp).ToString(); ;
            //        if (_y2Target != null)
            //        {
            //            tmpstring =
            //            ((double)dttest.Rows[i].ItemArray[1] > (double)dty2target.Rows[i].ItemArray[0]) ? tmpstring : "";
            //        }
            //        testitem = new Mtblib.Graph.Component.LabelPosition();
            //        testitem.Model = 2;
            //        testitem.Text = tmpstring;
            //        testitem.RowId = i + 1;
            //        testitem.FontColor = 4;
            //        _tsplot.DataLabel.PositionList.Add(testitem);
            //    }
            //    else
            //    {
            //        tmpstring = "";
            //        testitem = new Mtblib.Graph.Component.LabelPosition();

            //        testitem.Model = 2;
            //        testitem.Text = tmpstring;
            //        testitem.RowId = i + 1;
            //        _tsplot.DataLabel.PositionList.Add(testitem);
            //    }
            //}
            #endregion  

            #endregion

            //if (SymbolSize != null) _tsplot.Symbol.Size = SymbolSize;
            if (dcGroup.Count > 0)
            {
                Mtblib.Graph.Component.DataView.Symbol tmpSymbol = (Mtblib.Graph.Component.DataView.Symbol)_tsplot.Symbol.Clone();
                Mtblib.Graph.Component.DataView.Connect tmpConn = (Mtblib.Graph.Component.DataView.Connect)_tsplot.Connectline.Clone();
                tmpSymbol.GroupingBy = dcGroup.Values.ToArray();
                tmpSymbol.Color =
                    (_symbolcolor == 0) ? "lcolr" : _symbolcolor.ToString();
                tmpSymbol.Type = "lsymb";
                tmpSymbol.Visible = true;

                tmpConn.GroupingBy = dcGroup.Values.ToArray();
                tmpConn.Color = "lcolr";
                tmpConn.Type = "lline";
                tmpConn.Visible = true;
                
                tmpSymbol.Size = _symbolSize;
                tmpConn.Size = LineSize;
                tmpSymbol.Type = "ltype";

                foreach (Mtblib.Graph.Component.DataView.DataViewPosition itemtmp in dataViewPositionsList) tmpSymbol.DataViewPositionLst.Add(itemtmp);

                cmnd.Append(tmpSymbol.GetCommand());
                cmnd.Append(tmpConn.GetCommand());
            }
            else
            {
                Mtblib.Graph.Component.DataView.Symbol tmpSymbol = (Mtblib.Graph.Component.DataView.Symbol)_tsplot.Symbol.Clone();
                Mtblib.Graph.Component.DataView.Connect tmpConn = (Mtblib.Graph.Component.DataView.Connect)_tsplot.Connectline.Clone();

                tmpSymbol.Visible = true;
                tmpSymbol.Size = _symbolSize;
                tmpConn.Size = LineSize;

                tmpConn.Visible = true;
                cmnd.Append(tmpSymbol.GetCommand());
                cmnd.Append(tmpConn.GetCommand());
            }

            //設定 legend 要隱藏的 rows & columns (只有當 Group 和 XGroup 不同時需要處理)
            //使用 Dynamic LinQ 來處理
            if (dcGroup_color.Count > 0 && dcGroup.Count != dcGroup_color.Count)
            {
                #region Get the row number which need to be hided.
                var data = vars.Union(gps).ToArray();
                DataTable dt = Mtblib.Tools.MtbTools.GetDataTableFromMtbCols(gps);
                DataColumn dc = dt.Columns.Add("Variable", typeof(string));
                dc.SetOrdinal(0);

                DataTable dt_stacked = dt.Clone();
                DataRow dr;
                foreach (Mtb.Column col in vars) // Stack group data
                {
                    for (int r = 0; r < dt.Rows.Count; r++)
                    {
                        dr = dt.Rows[r];
                        dr[0] = col.Name;
                        dt_stacked.Rows.Add(dr.ItemArray);
                    }
                }
                string[] col_names = dt_stacked.Columns.Cast<DataColumn>().Select(x => x.ColumnName).ToArray();

                //var dt_group_by_gps = dt.AsEnumerable().GroupBy();
                string group_by_string = "new (" +
                    string.Join(",", col_names.Select(x => "it[\"" + x + "\"] as " + x)) + ")";

                col_names = new string[] { "Variable" };
                if (dcGroup_color.Count > 0)
                {
                    var tmp = col_names.ToList();
                    tmp.AddRange(gps_color.Select(x => x.SynthesizedName));
                    col_names = tmp.ToArray();
                }
                string select_string = "new (" + string.Join("+", col_names.Select(x => "Key." + x + ".ToString()")) + " as Uniq)";
                //string select_string = "new (" + string.Join(",", col_names.Select(x => "Key." + x + " as " + x)) + ")";

                var dt_legendbox = dt_stacked.AsEnumerable().GroupBy(group_by_string, "it").Select(select_string);
                var list_of_legendbox = (from dynamic g in dt_legendbox select g).ToList().Select(x => x.ToString()).ToList();


                group_by_string = "new (" + string.Join(",", col_names.Select(x => "it[\"" + x + "\"] as " + x)) + ")";

                var dt_legend_look = dt_stacked.AsEnumerable().GroupBy(group_by_string, "it").Select(select_string);
                var list_of_legendbox_look = (from dynamic g in dt_legend_look select g).ToList().Select(x => x.ToString()).ToList();

                List<int> row_index_to_keep = new List<int>();
                for (int i = 0; i < list_of_legendbox_look.Count; i++)
                {
                    var item = list_of_legendbox_look[i];
                    row_index_to_keep.Add(list_of_legendbox.IndexOf(item) + 1);
                }
                List<int> row_index = list_of_legendbox.Select((x, i) => i + 1).ToList();
                List<int> rhide = row_index.Except(row_index_to_keep).ToList();
                #endregion

                int start = dcVar.Count > 1 ? 3 : 2;
                int[] chide = gps.Select((x, i) => !gps_color.Contains(x) ? i + start : -1).Where(x => x > -1).ToArray();

                //取得 legend 本體
                List<Mtblib.Graph.Component.Region.LegendSection> mtb_sections = new List<Mtblib.Graph.Component.Region.LegendSection>();
                Mtblib.Graph.Component.Region.LegendSection section = new Mtblib.Graph.Component.Region.LegendSection(1);
                section.RowHide = rhide.ToArray();
                section.ColumnHide = chide;
                mtb_sections.Add(section);
                _tsplot.Legend.Sections = mtb_sections;
            }

            cmnd.Append(_tsplot.Legend.GetCommand());
            cmnd.Append(_tsplot.DataLabel.GetCommand());
            cmnd.Append(_tsplot.GetAnnotationCommand());
            cmnd.Append(_tsplot.GetOptionCommand());
            cmnd.Append(_tsplot.GetRegionCommand());
            cmnd.AppendLine(".");


            cmnd.AppendLine("endmacro");
            return cmnd.ToString();

        }

        public override void Run()
        {
            StringBuilder cmnd = new StringBuilder();
            string macroPath = Mtblib.Tools.MtbTools.BuildTemporaryMacro("mytrend.mac", GetCommand());
            cmnd.AppendLine("notitle");
            cmnd.AppendLine("brief 0");
            cmnd.AppendFormat("%\"{0}\" {1};\r\n", macroPath,
                string.Join(" &\r\n", ((Mtb.Column[])Variables).Select(x => x.SynthesizedName).ToArray()));
            if (GroupingBy != null)
            {
                cmnd.AppendFormat("group {0};\r\n", string.Join(" &\r\n",
                    ((Mtb.Column[])GroupingBy).Select(x => x.SynthesizedName).ToArray()));
            }
            if (_xgroups != null)
            {
                cmnd.AppendFormat("xgroup {0};\r\n", string.Join(" &\r\n",
                    ((Mtb.Column[])_xgroups).Select(x => x.SynthesizedName).ToArray()));
            }
            if (Stamp != null)
            {
                cmnd.AppendFormat("stamp {0};\r\n", string.Join(" &\r\n",
                    ((Mtb.Column[])Stamp).Select(x => x.SynthesizedName).ToArray()));
            }

            cmnd.AppendLine(".");
            cmnd.AppendLine("title");
            cmnd.AppendLine("brief 2");
            string fpath = Mtblib.Tools.MtbTools.BuildTemporaryMacro("mycode.mtb", cmnd.ToString());

            _proj.ExecuteCommand(string.Format("exec \"{0}\" 1", fpath));

        }

        
    }
}
