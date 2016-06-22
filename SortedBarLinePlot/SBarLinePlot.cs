using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.SortedBarLinePlot
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class SBarLinePlot : ISBarLinePlot, IDisposable
    {
        public SBarLinePlot()
        {

        }

        private Mtb.Project _proj;
        private Mtb.Worksheet _ws;
        private Mtblib.Graph.ScatterPlot.Plot _plot;
        private Mtblib.Graph.BarChart.Chart _chart;

        public void SetBarVariable(dynamic d)
        {
            BarVariable = d;
        }
        public void SetTrendVariable(dynamic d)
        {
            TrendVariable = d;
        }
        public void SetGroupingBy(dynamic d)
        {
            GroupingBy = d;
        }

        Mtb.Column[] _barvar = null;
        public dynamic BarVariable
        {
            get
            {
                return _barvar;
            }
            set
            {

                _barvar = Mtblib.Tools.MtbTools.GetMatchColumns(value, _ws);
            }
        }
        Mtb.Column[] _trndvar = null;
        public dynamic TrendVariable
        {
            get
            {
                return _trndvar;
            }
            set
            {
                _trndvar = Mtblib.Tools.MtbTools.GetMatchColumns(value, _ws);
            }
        }
        Mtb.Column[] _groupBy = null;
        public dynamic GroupingBy
        {
            get
            {
                return _groupBy;
            }
            set
            {
                _groupBy = Mtblib.Tools.MtbTools.GetMatchColumns(value, _ws);
            }
        }

        public Component.Scale.IContScale XScale
        {
            get
            {
                return new Component.Scale.Adapter_ContScale(_plot.XScale);
            }
        }

        public Component.Scale.IContScale YScale
        {
            get
            {
                return new Component.Scale.Adapter_ContScale(_chart.YScale);
            }
        }

        public Component.IDatlab Datlab
        {
            get
            {
                return new Component.Adapter_DatLab(_chart.DataLabel);
            }
        }

        public Component.ILabel Title
        {
            get
            {
                return new Component.Adapter_Lab(_chart.Title);
            }

        }

        public Component.Region.IRegion DataRegion
        {
            get
            {
                return new Component.Region.Adapter_Region(_chart.DataRegion);
            }
        }

        public Component.Region.ILegend Legend
        {
            get
            {
                return new Component.Region.Adapter_Legend(_chart.Legend);
            }
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
            _chart = new Mtblib.Graph.BarChart.Chart(_proj, _ws);
            _plot = new Mtblib.Graph.ScatterPlot.Plot(_proj, _ws);
            SetDefault();

        }
        public int TopK { set; get; }
        private void SetDefault()
        {
            /*
             * 因為使用疊圖法，為了顯示方便使用 Scatter plot 的 X軸、Bar chart 的 Y 軸讓使用者設定
             */
            TopK = 5;
            _plot.Connectline.Visible = true;
            _plot.YScale.LDisplay = new int[] { 1, 0, 0, 0 };
            _plot.YScale.HDisplay = new int[] { 1, 1, 1, 0 };
            _plot.DataRegion.Type = 0;
            _plot.DataRegion.EType = 0;
            _plot.Title.Visible = false;
            _plot.DataRegion.SetCoordinate(0.119, 0.8211, 0.1778, 0.9);
            _chart.DataRegion.SetCoordinate(0.119, 0.8211, 0.1778, 0.9);
            _chart.XScale.Ticks.HideAllTick = true;
            _chart.XScale.Label.Visible = false;
            _chart.BarsRepresent = Mtblib.Graph.BarChart.Chart.ChartRepresent.TWO_WAY_TABLE;
            _chart.YScale.GetCommand = () =>
            {
                #region Override GetCommand of YScale
                if (_chart.YScale.SecScale.Variable != null) throw new NotSupportedException("Bar Chart 不支援次座標變數");
                StringBuilder cmnd = new StringBuilder();
                if (_chart.YScale.LDisplay != null)
                    cmnd.AppendLine(string.Format("LDisplay {0};", string.Join(" ", _chart.YScale.LDisplay)));
                if (_chart.YScale.HDisplay != null)
                    cmnd.AppendLine(string.Format("HDisplay {0};", string.Join(" ", _chart.YScale.HDisplay)));
                if (_chart.YScale.Min < Mtblib.Tools.MtbTools.MISSINGVALUE) cmnd.AppendLine(string.Format("Min {0};", _chart.YScale.Min));
                if (_chart.YScale.Max < Mtblib.Tools.MtbTools.MISSINGVALUE) cmnd.AppendLine(string.Format("Max {0};", _chart.YScale.Max));

                cmnd.Append(_chart.YScale.Ticks.GetCommand());
                cmnd.Append(_chart.YScale.Label.GetCommand());
                if (cmnd.Length > 0) //如果有設定再加入
                    cmnd.Insert(0, string.Format("Scale {0};\r\n", (int)_chart.YScale.Direction));
                if (_chart.YScale.SecScale.Variable != null) cmnd.AppendLine("#SBarlineplot 不支援次座標變數設定 :(");
                return cmnd.ToString();
                #endregion
            };
        }
        /// <summary>
        /// 執行 Sbarlineplot 繪圖工具
        /// </summary>
        public void Run()
        {
            // 先檢查
            // 執行
            string fpath = Mtblib.Tools.MtbTools.BuildTemporaryMacro("chart.mac", GetCommand());
            StringBuilder cmnd = new StringBuilder();
            cmnd.AppendFormat("%\"{0}\" {1};\r\n", fpath,
                string.Join("&\r\n", ((Mtb.Column[])BarVariable).Select(x => x.SynthesizedName).ToArray()));
            cmnd.AppendFormat("trend {0};\r\n", ((Mtb.Column[])TrendVariable).Select(x => x.SynthesizedName).ToArray()[0]);
            cmnd.AppendFormat("group {0}.\r\n", ((Mtb.Column[])GroupingBy).Select(x => x.SynthesizedName).ToArray()[0]);
            fpath = Mtblib.Tools.MtbTools.BuildTemporaryMacro("mycode.mtb", cmnd.ToString());
            _proj.ExecuteCommand(string.Format("exec \"{0}\" 1", fpath), _ws);

        }

        /// <summary>
        /// 取得 SBar-line plot 的指令碼
        /// </summary>
        /// <returns></returns>
        private string GetCommand()
        {
            Mtb.Column[] barvar;
            Mtb.Column[] trndvar;
            if (BarVariable == null || TrendVariable == null)
            {
                throw new ArgumentNullException("Bar variable 和 Trend variable 不可為空");
            }
            else
            {
                barvar = (Mtb.Column[])BarVariable;
                trndvar = (Mtb.Column[])TrendVariable;
            }

            if (GroupingBy == null) throw new ArgumentNullException("分群資訊不可為空");

            Mtb.Column[] gps = (Mtb.Column[])GroupingBy;
            if (gps.Length > 1) throw new ArgumentException("只能指定一個分群欄位");
            if (gps[0].RowCount != gps[0].GetNumDistinctRows()) throw new ArgumentException("分群欄位中有重複的項目");

            dynamic values = gps[0].GetData();

            string[] gp = null;
            if (values is double[])
            {
                gp = ((double[])values).Select(x => x.ToString()).ToArray();
            }
            else
            {
                gp = (string[])values;
            }


            var gpInfo = gp.Select((x, i) => new { Name = x, Order = i + 1 });


            StringBuilder cmnd = new StringBuilder();
            cmnd.AppendLine("macro");
            cmnd.AppendLine("sbarline y.1-y.n;");
            cmnd.AppendLine("trend trnd;");
            cmnd.AppendLine("group x.");
            cmnd.AppendLine("mcolumn y.1-y.n trnd x");
            cmnd.AppendLine("mcolumn xx yy ylab txx ttrnd rank");
            cmnd.AppendLine("mconstant symax");

            cmnd.AppendLine("mreset");
            cmnd.AppendLine("notitle");
            cmnd.AppendLine("noecho");
            cmnd.AppendLine("brief 0");

            cmnd.AppendLine("stack y.1-y.n yy;");
            cmnd.AppendLine("subs ylab;");
            cmnd.AppendLine("usen.");
            cmnd.AppendLine("tset xx");
            cmnd.AppendFormat("{1}({0})1\r\n",
                string.Join("&\r\n", gp.Select(x => "\"" + x + "\"").ToArray()),
                barvar.Count());
            cmnd.AppendLine("end");

            #region 用 C# 計算分組 rank 值
            List<double> stackedBarValues = new List<double>();
            List<string> stackedGroupName = new List<string>();
            List<string> stackedSubscript = new List<string>();
            for (int i = 0; i < barvar.Length; i++)
            {
                for (int j = 0; j < barvar[i].RowCount; j++)
                {
                    stackedSubscript.Add(barvar[i].Name);
                }
                stackedGroupName.AddRange(gp);
                stackedBarValues.AddRange(barvar[i].GetData());
            }
            double[] rank = Mtblib.Tools.MtbTools.GroupRank(
                stackedBarValues.ToArray(),
                stackedGroupName.ToArray(),
                true,
                Mtblib.Tools.MtbTools.RankType.RANK);
            #endregion

            cmnd.AppendLine("set rank");
            cmnd.AppendLine(string.Join(" &\r\n", rank));
            cmnd.AppendLine("end");
            cmnd.AppendLine("copy xx ylab yy xx ylab yy;");
            cmnd.AppendLine(" include;");
            cmnd.AppendFormat(" where \"rank<={0}\".\r\n", TopK);

            cmnd.AppendLine("copy trnd ttrnd");


            #region 用 C# 處理虛擬欄位內容
            //取得各分群中，前K大的項目。這裡不使用 Minitab macro，因為可能會很慢
            var topKgpName = stackedGroupName.ToArray().Zip(rank, (x, r) => new
            {
                Name = x,
                Rank = r
            }).Zip(stackedSubscript, (x, i) => new
            {
                Name = x.Name,
                Rank = x.Rank,
                Item = i
            })
            .Where(x => x.Rank <= TopK);

            /* 
             * 計算各組共取得多少前K大項目，因為有可能超過K的(因為 tied rank)
             * 該匿名類型只是過客..下一個 linq 的產出才是主角 XDDD
             */
            var xx = from x in
                         (from item in topKgpName
                          group item by item.Name into g
                          select new { g.Key, Count = g.Count() })
                     join y in gpInfo on x.Key equals y.Name into xy
                     from z in xy
                     select new
                     {
                         Name = z.Name,
                         Count = x.Count,
                         Order = z.Order,
                     };
            /*
             * 計算各組虛擬項目與值的資訊(有多少要排在前面、多少排在後面)，
             * 因為 Group by 之後的分群順序不一定是使用者指定順序
             * 
             */
            var countEachGroupInfo
                = from x in xx
                  select new
                  {
                      Name = x.Name,
                      Count = x.Count,
                      Order = x.Order,
                      Zero = (from o in xx
                              where o.Order > x.Order
                              select o.Count).Sum(),
                      Missing = (from o in xx
                                 where o.Order < x.Order
                                 select o.Count).Sum()
                  };


            //圖中 Bar 的數量(包含間距)
            int ttlBarCount = countEachGroupInfo.Select(x => x.Count).ToArray().Sum() +
                gpInfo.Count() - 1;


            List<string> dummyGpName = new List<string>();
            List<string> dummyItems = new List<string>();
            List<double> dummyValues = new List<double>();
            List<double> dummyX = new List<double>();// Trend 的 X 座標
            // 將要繪製的 item 抓出，並依據字母順序給定 order
            Dictionary<string, int> allItems = topKgpName.Select(x => x.Item).Distinct().OrderBy(x=>x).
                Select((x, i) => new { Item = x, Index = i }).ToDictionary(x => x.Item, x => x.Index + 1);
            for (int i = 0; i < gpInfo.Count(); i++)
            {
                string name = gpInfo.Select(x => x.Name).ToArray()[i];
                int zeroCount = countEachGroupInfo.Where(x => x.Name == name).Select(x => x.Zero).First() + gpInfo.Count() - i - 1;
                int missingCount = countEachGroupInfo.Where(x => x.Name == name).Select(x => x.Missing).First() + i;
                int dummyCount = zeroCount + missingCount;

                #region 判斷某群組內是否有應包含而未包含的 Bar Item
                /*
                 * 如果有一些 bar item 未在此群組內，那麼最後圖形的 legend box 顯示會不正確
                 * e.g. 全部群組要畫的 bar: a, b, c, d, e。因為用第一張 bar chart 的 legend box
                 * 做為最終圖形的 legend box，如果群組1只包含 a,b,c 而未加入 d, e；bar 的顏色將與
                 * 其他圖形不一致，且 legend box 顯示錯誤。
                 */
                var barItem = from b in topKgpName
                              where b.Name == name
                              select b.Item;
                var nonContainedItem = allItems.Select(x => x.Key).Except(barItem);
                int nonContainedItemCount = 0;
                if (nonContainedItem.Any())
                {
                    nonContainedItemCount = nonContainedItem.Count();

                }
                #endregion

                for (int j = 0; j < dummyCount; j++)
                {
                    dummyGpName.Add(name);
                    if (j <= nonContainedItemCount - 1)
                    {
                        dummyItems.Add(nonContainedItem.ToArray()[j]);
                    }
                    else
                    {
                        dummyItems.Add("ZZZZZZZ" + j.ToString("D3"));
                    }

                    if (j < zeroCount)
                    {
                        dummyValues.Add(0);
                    }
                    else
                    {
                        dummyValues.Add(Mtblib.Tools.MtbTools.MISSINGVALUE);
                    }
                }
                dummyX.Add((ttlBarCount - zeroCount + missingCount + 1) / 2);
            }





            #endregion

            /*
             * 加入虛擬欄位資訊於 macro 中
             * dummayGpName, dummyItems, dummayValues, dummyX
             */
            cmnd.AppendLine("insert xx");
            cmnd.AppendFormat("{0}\r\n", string.Join(" &\r\n", dummyGpName.Select(x => "\"" + x + "\"").ToArray()));
            cmnd.AppendLine("end");
            cmnd.AppendLine("insert ylab");
            cmnd.AppendFormat("{0}\r\n", string.Join(" &\r\n", dummyItems.Select(x => "\"" + x + "\"").ToArray()));
            cmnd.AppendLine("end");
            cmnd.AppendLine("insert yy");
            cmnd.AppendFormat("{0}\r\n", string.Join(" &\r\n", dummyValues));
            cmnd.AppendLine("end");
            cmnd.AppendLine("set txx");
            cmnd.AppendLine(string.Join("&\r\n", dummyX));
            cmnd.AppendLine("end");

            /*
             * 計算所有 bar 和 trend 的 scale 資訊
             * Bar 的資訊用來設定疊圖座標軸的上限值
             * Trend 的資訊會用於調整 Data region 大小
             * 
             */
            List<double[]> allvalues = new List<double[]>();
            allvalues.Add(stackedBarValues.ToArray());
            allvalues.Add(trndvar[0].GetData());
            Mtblib.Tools.GScale[] gscale = Mtblib.Tools.MtbTools.GetMinitabGScaleInfo(allvalues, _proj, _ws);
            _chart.YScale.Max = gscale[0].SMax;
            _chart.XScale.Ticks.HideAllTick = true;
            _chart.Bar.GroupingBy = "ylab";

            /*
             * 計算 legend box 寬 + 次座標 Axlab 高 + 次座標 Tick 寬的像素
             * 預設為 0.119, 0.8211,0.1778,0.9，如果使用者指定這組將依預設方式處理。
             * 
             */
            #region 調整繪圖區座標

            if (_chart.DataRegion.GetCoordinate().SequenceEqual(new double[] { 0.119, 0.8211, 0.1778, 0.9 }))
            {
                #region 預設調整
                double w = 0;
                float fsize = 0;
                string txt = null;

                fsize = fsize = _chart.Legend.FontSize > 0 ? _chart.Legend.FontSize : 7f;

                w = w + Mtblib.Tools.MtbTools.GetSizeOfString(
                    allItems.Select(x => x.Key).ToArray(), new System.Drawing.Font("Segoe UI Semibold", fsize)
                    ).Select(x => x.Width).Max() + 28;

                txt = string.IsNullOrEmpty(_chart.YScale.SecScale.Label.Text) ? trndvar[0].Name : _chart.YScale.SecScale.Label.Text;
                fsize = _chart.YScale.SecScale.Label.FontSize > 0 ? _chart.YScale.SecScale.Label.FontSize : 11f;
                w = w + Mtblib.Tools.MtbTools.GetSizeOfString(txt, new System.Drawing.Font("Segoe UI Semibold", fsize))[0].Height;

                fsize = _chart.YScale.SecScale.Ticks.FontSize > 0 ? _chart.YScale.SecScale.Ticks.FontSize : 8f;
                w = w + Mtblib.Tools.MtbTools.GetSizeOfString(
                    new string[] { gscale[1].TMaximum.ToString(), gscale[1].TMinimum.ToString() },
                    new System.Drawing.Font("Segoe UI Semibold", fsize)).Select(x => x.Width).Max();

                float ww = 115;
                if (w > ww)
                {
                    double[] coord = _chart.DataRegion.GetCoordinate();
                    _chart.DataRegion.SetCoordinate(
                        coord[0], coord[1] + (-0.004) * (w - ww), coord[2], coord[3]
                        );
                    _plot.DataRegion.SetCoordinate(
                        coord[0], coord[1] + (-0.004) * (w - ww), coord[2], coord[3]);
                }
                else
                {
                    _chart.DataRegion.SetCoordinate(0.119, 0.8211, 0.1778, 0.9);
                }
                #endregion
            }
            else//表示有更動過
            {
                double[] coord = _chart.DataRegion.GetCoordinate();
                _plot.DataRegion.SetCoordinate(coord[0], coord[1], coord[2], coord[3]);
            }

            #endregion

            /*
             * 確認有多少 item 要顯示在 legend box 中
             */
            int legendItems = topKgpName.Select(x => x.Item).Distinct().Count();

            cmnd.AppendLine("layout;");
            cmnd.Append(_chart.GetOptionCommand());
            cmnd.AppendLine(".");

            #region 建立 Bar 指令

            /*
             * 過程中會變動 Chart 物件的屬性，所以畫完後要還原，建立備份物件
             */
            Mtblib.Graph.Component.Scale.ContScale _yscale =
                (Mtblib.Graph.Component.Scale.ContScale)_chart.YScale.Clone();
            Mtblib.Graph.Component.Title _title =
                (Mtblib.Graph.Component.Title)_chart.Title.Clone();

            /*
             * 設定的 Datlab 隱藏資訊
             * 
             * 版本一: 用 one way table 去畫 (取消..因為 legend box 順序問題)
             * 將如果要顯示 Datlab，計算出要隱藏的 position，
             * 對 summarized data 來說，即要找出來源資料不想顯示的 row id
             * 
             * 假設有 k 組，第 i 組要繪製的 bar 有 ni 個，且 n1+...nk=N，加上虛擬 bar = N+(k-1)=M
             * 對第 i 組要隱藏的 row id = (i-1)*M+N-(n1+..n(i-1))+1 ~ i*M+N-(n1+..+ni)
             * 因此，需要一些 summarize info 幫助計算
             * M = ttlBarCount
             * N = countEachGroupInfo 中 Count 的和
             * ni = countEachGroupInfo 中第 i 組 Count
             * 
             * 版本二: 用 function of variable 去畫
             * 直接對圖形中要隱藏的 bar 的位置，位置依據 legend box 的順序決定，
             * 所以要知道哪一些 item 不顯示
             * 
             */
            int baseBarCount = countEachGroupInfo.Select(x => x.Count).ToArray().Sum();

            for (int i = 0; i < gp.Length; i++)
            {
                cmnd.AppendFormat("##### chart for \"{0}\" #####\r\n", gp[i]);
                cmnd.AppendLine("chart sum(yy)*ylab;");
                //cmnd.AppendLine("summ;");
                if (i == 1)
                {
                    _chart.YScale.LDisplay = new int[] { 0, 0, 0, 0 };
                    _chart.YScale.Ticks.HideAllTick = true;
                    _chart.YScale.Label.Visible = false;
                    _chart.DataRegion.Type = 0;
                    _chart.DataRegion.EType = 0;
                    _chart.Title.Visible = false;
                    _chart.Legend.HideLegend = true;
                }
                if (_chart.YScale.Label.Text == null) _chart.YScale.Label.Text = "Count";
                cmnd.Append(_chart.YScale.GetCommand());
                cmnd.Append(_chart.XScale.GetCommand());
                cmnd.Append(_chart.Bar.GetCommand());

                cmnd.AppendLine("gapw 0;");
                cmnd.AppendLine("decr;");
                Mtblib.Graph.Component.Region.LegendSection lsection
                    = new Mtblib.Graph.Component.Region.LegendSection(1);
                lsection.HideColumnHeader = true;
                lsection.RowHide = Enumerable.Range(legendItems + 1, ttlBarCount+1).ToArray();
                _chart.Legend.Sections.Add(lsection);
                cmnd.Append(_chart.Legend.GetCommand());
                //if (i == 0)
                //{
                //    cmnd.AppendLine("legend;");
                //    cmnd.AppendLine("sect 1;");
                //    cmnd.AppendLine("chhide;");
                //    cmnd.AppendFormat("rhide {0}:{1};\r\n", legendItems + 1, ttlBarCount);
                //}
                //else
                //{
                //    cmnd.AppendLine("nolegend;");
                //}
                cmnd.AppendLine("includ;");
                cmnd.AppendFormat("where \"xx=\"\"{0}\"\"\";\r\n", gp[i]);


                #region 版本一的取 Position 方法(已關閉)
                //// 要隱藏的 Position 起點(版本一)
                //int start = (i) * ttlBarCount + baseBarCount - cumulativeCount + 1;
                //// 取出第 i 組的組數
                //cumulativeCount = cumulativeCount +
                //    countEachGroupInfo.Where(x => x.Name == gp[i]).Select(x => x.Count).First();
                //// 要隱藏的 Position 終點
                //int end = (i + 1) * ttlBarCount + baseBarCount - cumulativeCount;

                //_chart.DataLabel.PosititionList = new List<Mtblib.Graph.Component.LabelPosition>(); //清空設定
                //Mtblib.Graph.Component.LabelPosition pos; //暫存用的 pos 類別
                //for (int j = start; j <= end; j++)
                //{
                //    pos = new Mtblib.Graph.Component.LabelPosition(j, "");
                //    _chart.DataLabel.PosititionList.Add(pos);
                //} 
                #endregion

                var barItem = from b in topKgpName
                              where b.Name == gp[i]
                              select b.Item;
                var nonContainedItem = allItems.Select(x => x.Key).Except(barItem);
                if (nonContainedItem.Any())
                {
                    foreach (var item in nonContainedItem)
                    {
                        _chart.DataLabel.PositionList.Add(
                            new Mtblib.Graph.Component.LabelPosition(allItems[item], ""));
                    }
                }
                for (int j = allItems.Count + 1; j <= ttlBarCount; j++)
                {
                    _chart.DataLabel.PositionList.Add(
                            new Mtblib.Graph.Component.LabelPosition(j, ""));
                }

                cmnd.Append(_chart.DataLabel.GetCommand());
                _chart.DataLabel.PositionList = new List<Mtblib.Graph.Component.LabelPosition>(); //清空設定，下一個會不同，且沒有開放介面

                cmnd.Append(_chart.DataRegion.GetCommand());
                if (_chart.Title.Text == null) _chart.Title.Text = "Bar line plot";
                cmnd.Append(_chart.Title.GetCommand());
                cmnd.Append(_chart.FigureRegion.GetCommand());
                cmnd.AppendLine("nosf;");
                cmnd.AppendLine("noxf.");
            }

            //還原使用者的設定
            _chart.YScale = (Mtblib.Graph.Component.Scale.ContScale)_yscale.Clone();
            _chart.Title = (Mtblib.Graph.Component.Title)_title.Clone();
            _chart.Legend.HideLegend = false;
            _chart.Legend.Sections = new List<Mtblib.Graph.Component.Region.LegendSection>();
            _chart.DataRegion.Type = null;
            _chart.DataRegion.EType = null;

            #endregion


            #region 建立 Trend 指令
            Mtblib.Graph.Component.Scale.AxLabel axlab1 = (Mtblib.Graph.Component.Scale.AxLabel)_plot.XScale.Label.Clone();
            Mtblib.Graph.Component.Scale.AxLabel axlab2 = (Mtblib.Graph.Component.Scale.AxLabel)_chart.YScale.SecScale.Label.Clone();
            axlab2.Side = 2;
            axlab2.ScalePrimary = Mtblib.Graph.Component.ScalePrimary.Primary;
            _plot.YScale.Label = axlab2;
            _plot.YScale.Ticks = (Mtblib.Graph.Component.Scale.ContTick)_chart.YScale.SecScale.Ticks.Clone();
            _plot.YScale.Min = _chart.YScale.SecScale.Min;
            _plot.YScale.Max = _chart.YScale.SecScale.Max;
            _plot.XScale.Min = 0.5;
            _plot.XScale.Max = 0.5 + ttlBarCount;
            _plot.XScale.Ticks.SetTicks("txx");
            _plot.XScale.Ticks.SetLabels(gp.Select(x => "\"" + x + "\"").ToArray());
            _plot.DataLabel = (Mtblib.Graph.Component.Datlab)_chart.DataLabel.Clone();
            _plot.Symbol.Visible = true;
            _plot.Connectline.Visible = true;
            cmnd.AppendFormat("##### scatter plot #####\r\n");
            cmnd.AppendLine("plot ttrnd*txx;");
            if (_plot.YScale.Label.Text == null) _plot.YScale.Label.Text = trndvar[0].Name;
            cmnd.Append(_plot.YScale.GetCommand());

            if (_plot.XScale.Label.Text == null) _plot.XScale.Label.Text = gps[0].Name;
            cmnd.Append(_plot.XScale.GetCommand());
            cmnd.Append(_plot.Symbol.GetCommand());
            cmnd.Append(_plot.Connectline.GetCommand());
            cmnd.Append(_plot.DataLabel.GetCommand());
            cmnd.Append(_plot.DataRegion.GetCommand());
            cmnd.Append(_plot.FigureRegion.GetCommand());
            cmnd.Append(_plot.Title.GetCommand());
            cmnd.AppendLine("nosf;");
            cmnd.AppendLine("noxf.");

            //還原修改過的元件
            _plot.XScale.Label = (Mtblib.Graph.Component.Scale.AxLabel)axlab1.Clone();
            _plot.YScale.Label = (Mtblib.Graph.Component.Scale.AxLabel)axlab2.Clone();

            #endregion

            cmnd.AppendLine("endlayout");

            cmnd.AppendLine("endmacro");

            //由 Mtb.Column 取得 distinct 的值
            //對每一個群組執行排序


            return cmnd.ToString();
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
                _plot.Dispose();
            }
            // Free your own state (unmanaged objects).
            // Set large fields to null.
            _proj = null;
            _ws = null;
            _barvar = null;
            _trndvar = null;
            _groupBy = null;
            GC.Collect();

        }
        ~SBarLinePlot()
        {
            Dispose(false);
        }
    }
}
