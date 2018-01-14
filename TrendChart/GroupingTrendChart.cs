﻿using System;
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
        }

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


            StringBuilder cmnd = new StringBuilder();

            cmnd.AppendLine("macro");
            cmnd.AppendLine("mytrend y.1-y.ny;");
            cmnd.AppendLine("groupby g;");
            cmnd.AppendLine("xgroup x.1-x.m;");
            cmnd.AppendLine("stamp stmp");

            cmnd.AppendLine("mcolumn y.1-y.ny ");
            cmnd.AppendLine("mcolumn x.1-x.m g stmp");
            cmnd.AppendLine("mcolumn xx.1-xx.4 nn.1-nn.4 yy uniq freq index link vline");
            cmnd.AppendLine("mcolumn colr symb line lcolr lsymb lline tmp.1-tmp.10");
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
                cmnd.AppendFormat(" cnt({0})\r\n", string.Join(" ", Mtblib.Tools.MtbTools.COLOR));
                cmnd.AppendLine("end");
                cmnd.AppendLine("let const.2 = Count(colr)");
                cmnd.AppendLine("delete const.1:const.2 colr");
                //Set symbol list 
                cmnd.AppendLine("set symb");
                cmnd.AppendFormat(" cnt({0})\r\n", string.Join(" ", Mtblib.Tools.MtbTools.SYMBOLTYPE));
                cmnd.AppendLine("end");
                cmnd.AppendLine("let const.2 = Count(symb)");
                cmnd.AppendLine("delete const.1:const.2 symb");
                //Set symbol list 
                cmnd.AppendLine("set line");
                cmnd.AppendFormat(" cnt({0})\r\n", string.Join(" ", Mtblib.Tools.MtbTools.LINETPYE));
                cmnd.AppendLine("end");
                cmnd.AppendLine("let const.2 = Count(line)");
                cmnd.AppendLine("delete const.1:const.2 line");
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
            if (dcGroup_x.Count > 0)
            {
                _tsplot.XScale.Refes.Values = "vline";
                _tsplot.XScale.Refes.Color = new string[] { "20" };
            }
            cmnd.Append(_tsplot.XScale.GetCommand());
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

            
            if (dcGroup.Count > 0)
            {
                Mtblib.Graph.Component.DataView.Symbol tmpSymbol = (Mtblib.Graph.Component.DataView.Symbol)_tsplot.Symbol.Clone();
                Mtblib.Graph.Component.DataView.Connect tmpConn = (Mtblib.Graph.Component.DataView.Connect)_tsplot.Connectline.Clone();
                tmpSymbol.GroupingBy = dcGroup.Values.ToArray();
                tmpSymbol.Color = "lcolr";
                tmpSymbol.Type = "lsymb";
                tmpSymbol.Visible = true;

                tmpConn.GroupingBy = dcGroup.Values.ToArray();
                tmpConn.Color = "lcolr";
                tmpConn.Type = "lline";
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