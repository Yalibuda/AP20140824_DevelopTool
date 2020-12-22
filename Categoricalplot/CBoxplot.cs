using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using MtbGraph.Component;
using MtbGraph.Component.DataView;
using MtbGraph.Component.Region;
using MtbGraph.Component.Scale;

namespace MtbGraph.Categoricalplot
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class CBoxplot : ICBoxplot, IDisposable
    {
        // Composite Boxplot
        public CBoxplot()
        {

        }
        public CBoxplot(Mtb.Project proj, Mtb.Worksheet ws)
        {
            SetMtbEnvironment(proj, ws);
        }
        protected Mtb.Project _proj;
        protected Mtb.Worksheet _ws;
        protected Mtblib.Graph.CategoricalChart.BoxPlot _boxplot;

        public void Run()
        {
            string cmnd = GetCommand();
            string fpath = Mtblib.Tools.MtbTools.BuildTemporaryMacro("mycode.mtb", cmnd.ToString());
            _proj.ExecuteCommand(string.Format("exec \"{0}\" 1", fpath));
        }

        protected virtual string GetCommand()
        {
            Mtb.Column[] vars;
            if (Variables == null) throw new ArgumentNullException("Variable 不可為空");
            else vars = Variables;

            Mtb.Column[] gps = GroupingVariables;

            StringBuilder cmnd = new StringBuilder();

            #region compute stats
            Mtb.Column[] columnname;
            Mtb.Column[] columnmin = null;
            Mtb.Column[] columnmedian = null;
            Mtb.Column[] columnq1 = null;
            Mtb.Column[] columnmean = null;
            Mtb.Column[] columnq3 = null;
            Mtb.Column[] columnmax = null;
            Mtb.Column[] columncount = null;
            int counttmp = 501;
            int countstats = 0;

            cmnd.AppendFormat("Name C{0} \"Stats\" ", (counttmp - 1).ToString());
            if (MinVisible)
            {
                cmnd.AppendFormat("C{0} \"Stats_Min\" ", counttmp);
                counttmp++;
                countstats++;
            }
            if (Q1Visible)
            {
                cmnd.AppendFormat("C{0} \"Stats_Q1\" ", counttmp);
                counttmp++;
                countstats++;
            }
            if (MedianVisible)
            {
                cmnd.AppendFormat("C{0} \"Stats_Median\" ", counttmp);
                counttmp++;
                countstats++;
            }
            if (MeanVisible)
            {
                cmnd.AppendFormat("C{0} \"Stats_Mean\" ", counttmp);
                counttmp++;
                countstats++;
            }
            if (Q3Visible)
            {
                cmnd.AppendFormat("C{0} \"Stats_Q3\" ", counttmp);
                counttmp++;
                countstats++;
            }
            if (MaxVisible)
            {
                cmnd.AppendFormat("C{0} \"Stats_Max\" ", counttmp);
                counttmp++;
                countstats++;
            }
            if (CountVisible)
            {
                cmnd.AppendFormat("C{0} \"Stats_Count\" ", counttmp);
                counttmp++;
                countstats++;
            }
            cmnd.AppendLine(".");

            cmnd.AppendFormat("Statistics {0}; \r\n", string.Join(" &\r\n", vars.Select(x => x.SynthesizedName).ToArray()));
            if (gps != null)
            {
                cmnd.AppendFormat(" By {0} ; \r\n", gps[0].SynthesizedName);
                cmnd.AppendFormat("  GValues 'Stats'; \r\n");
            }
            else
            {

            }
            if (MinVisible) cmnd.AppendLine(" Minimum 'Stats_Min';");
            if (Q1Visible) cmnd.AppendLine(" QOne 'Stats_Q1';");
            if (MedianVisible) cmnd.AppendLine(" Median 'Stats_Median';");
            if (MeanVisible) cmnd.AppendLine("  Mean 'Stats_Mean';");
            if (Q3Visible) cmnd.AppendLine(" QThree 'Stats_Q3';");
            if (MaxVisible) cmnd.AppendLine(" Maximum 'Stats_Max';");
            if (CountVisible) cmnd.AppendLine(" Count 'Stats_Count';");
            cmnd.AppendLine(".");

            string fpath = Mtblib.Tools.MtbTools.BuildTemporaryMacro("mycode.mtb", cmnd.ToString());
            _proj.ExecuteCommand(string.Format("exec \"{0}\" 1", fpath));

            cmnd.Clear();

            // number of stats

            dynamic t = Mtblib.Tools.MtbTools.GetMatchColumns("Stats", _ws);
            columnname = t;
            dynamic tmin;
            if (MinVisible)
            {
                tmin = Mtblib.Tools.MtbTools.GetMatchColumns("Stats_Min", _ws);
                columnmin = tmin;
            }
            dynamic tq1;
            if (Q1Visible)
            {
                tq1 = Mtblib.Tools.MtbTools.GetMatchColumns("Stats_Q1", _ws);
                columnq1 = tq1;
            }
            
            dynamic tmedian;
            if (MedianVisible)
            {
                tmedian = Mtblib.Tools.MtbTools.GetMatchColumns("Stats_Median", _ws);
                columnmedian = tmedian;
            }
            
            dynamic tmean;
            if (MeanVisible)
            {
                tmean = Mtblib.Tools.MtbTools.GetMatchColumns("Stats_Mean", _ws);
                columnmean = tmean;
            }

            dynamic tq3;
            if (Q3Visible)
            {
                tq3 = Mtblib.Tools.MtbTools.GetMatchColumns("Stats_Q3", _ws);
                columnq3 = tq3;
            }
           
            dynamic tmax;
            if (MaxVisible)
            {
                tmax = Mtblib.Tools.MtbTools.GetMatchColumns("Stats_Max", _ws);
                columnmax = tmax;
            }
            dynamic tcount;
            if (CountVisible)
            {
                tcount = Mtblib.Tools.MtbTools.GetMatchColumns("Stats_Count", _ws);
                columncount = tcount;
            }
            

            //for (int i = 0; i < colummean[0].RowCount; i++)
            //{
            //    double aaa = colummean[0].GetData(i + 1);
            //    Console.WriteLine("{0}", aaa);
            //}
            #endregion

            #region Create Boxplot
            if (gps != null)
            {
                cmnd.AppendFormat("Boxplot ({0})*{1};\r\n",
                    string.Join(" &\r\n", vars.Select(x => x.SynthesizedName).ToArray()),
                    gps[0].SynthesizedName);
                if (gps.Length >= 2)
                    throw new ArgumentException(
                        string.Format("輸入分群數 {0}，不是合法的數量。CBoxplot 用於分群的欄位數量最多為 1。", gps.Length));
                //cmnd.AppendFormat(" Group {0};\r\n",
                //string.Join(" &\r\n", gps.Select((x, i) => new { colId = x.SynthesizedName, index = i }).
                //Where(x => x.index > 0).Select(x => x.colId).ToArray()));
            }
            else
            {
                cmnd.AppendFormat("Boxplot {0};\r\n",
                    string.Join(" &\r\n", vars.Select(x => x.SynthesizedName).ToArray()));
            }

            cmnd.Append(_boxplot.GetOptionCommand());

            if (GSave != null)
            {
                _boxplot.GraphPath = GSave;
                cmnd.Append(_boxplot.GetOptionCommand());
                _boxplot.GraphPath = null;
            }

            cmnd.Append(_boxplot.YScale.GetCommand());
            cmnd.Append(_boxplot.XScale.GetCommand());

            cmnd.Append(_boxplot.Mean.GetCommand());
            cmnd.Append(_boxplot.CMean.GetCommand());
            cmnd.Append(_boxplot.RBox.GetCommand());
            cmnd.Append(_boxplot.IQRBox.GetCommand());
            cmnd.Append(_boxplot.Whisker.GetCommand());
            cmnd.Append(_boxplot.Outlier.GetCommand());
            //cmnd.Append(_boxplot.Title.GetCommand());

            cmnd.Append(_boxplot.Individual.GetCommand());
            cmnd.Append(_boxplot.MeanDatlab.GetCommand());
            cmnd.Append(_boxplot.IndivDatlab.GetCommand());

            cmnd.Append(_boxplot.Panel.GetCommand());

            // Prepare text box
            double yminregion = 0.2925; // 如果不在用到 可用 t=t+a
            _division = (IfStatVisible) ? yminregion + countstats * 0.03 : yminregion;
            double xmindataregion = 0.1065;
            double xmaxdataregion = 0.9533;
            _boxplot.DataRegion.AutoSize = false;
            _boxplot.DataRegion.SetCoordinate(xmindataregion, xmaxdataregion, _division, 0.8814);
            cmnd.Append(_boxplot.GetRegionCommand());
                        
            if (IfStatVisible)
            {
                double xgroupunit = (xmaxdataregion - xmindataregion) / columnname[0].RowCount;
                //先顯示總共要多少個統計量於左側
                Mtblib.Graph.Component.Annotation.Textbox tb = new Mtblib.Graph.Component.Annotation.Textbox();
                MtbGraph.Component.Annotation.ITextBox textBox = new MtbGraph.Component.Annotation.Adapter_Textbox(tb);
                textBox.SetCoordinate(xmindataregion - 3 * xgroupunit / 2, yminregion); // 用group總個數算x
                string statstext = "";
                if (MinVisible) statstext = statstext + "Min";
                if (Q1Visible) statstext = (statstext == "") ? ("Q1") : (statstext + "\\rQ1");
                if (MedianVisible) statstext = (statstext == "") ? ("Median") : (statstext + "\\rMedian");
                if (MeanVisible) statstext = (statstext == "") ? ("Mean") : (statstext + "\\rMean");
                if (Q3Visible) statstext = (statstext == "") ? ("Q3") : (statstext + "\\rQ3");
                if (MaxVisible) statstext = (statstext == "") ? ("Max") : (statstext + "\\rMax");
                if (CountVisible) statstext = (statstext == "") ? ("Count") : (statstext + "\\rCount");
                textBox.Text = "\"" + statstext + "\"";
                textBox.Unit = 0;
                //textBox.SetTextSize(TextSize);

                
                textBox.SetBoxposition(xmindataregion - 3 * xgroupunit / 2,
                    0.99, 0, 0.08 + countstats * 0.03);
                cmnd.AppendLine(textBox.GetCommand());

                // 整理每一群的各個統計量
                for (int i = 0; i < columnname[0].RowCount; i++)
                {
                    string tttt = "";
                    tb = new Mtblib.Graph.Component.Annotation.Textbox();
                    textBox = new MtbGraph.Component.Annotation.Adapter_Textbox(tb);
                    textBox.SetCoordinate(xmindataregion + xgroupunit * (i + 1 / 2), 0); // 用group總個數算x
                    string tmptext = "";
                    if (MinVisible)
                    {
                        tmptext = columnmin[0].GetData(i + 1).ToString();
                        tmptext = (tmptext == Mtblib.Tools.MtbTools.MISSINGVALUE.ToString()) ? "*" : tmptext;
                        tttt = (tttt == "") ? (tmptext) : (tttt + "\\r" + tmptext);
                    }
                    if (Q1Visible)
                    {
                        tmptext = columnq1[0].GetData(i + 1).ToString();
                        tmptext = (tmptext == Mtblib.Tools.MtbTools.MISSINGVALUE.ToString()) ? "*" : tmptext;
                        tttt = (tttt == "") ? (tmptext) : (tttt + "\\r" + tmptext); 
                    }
                    if (MedianVisible)
                    {
                        tmptext = columnmedian[0].GetData(i + 1).ToString();
                        tmptext = (tmptext == Mtblib.Tools.MtbTools.MISSINGVALUE.ToString()) ? "*" : tmptext;
                        tttt = (tttt == "") ? (tmptext) : (tttt + "\\r" + tmptext);
                    }
                    if (MeanVisible)
                    {
                        tmptext = Math.Round((double)columnmean[0].GetData(i + 1),2).ToString();
                        tmptext = (tmptext == Mtblib.Tools.MtbTools.MISSINGVALUE.ToString()) ? "*" : tmptext;
                        tttt = (tttt == "") ? (tmptext) : (tttt + "\\r" + tmptext);
                    }
                    if (Q3Visible)
                    {
                        tmptext = columnq3[0].GetData(i + 1).ToString();
                        tmptext = (tmptext == Mtblib.Tools.MtbTools.MISSINGVALUE.ToString()) ? "*" : tmptext;
                        tttt = (tttt == "") ? (tmptext) : (tttt + "\\r" + tmptext);
                    }
                    if (MaxVisible)
                    {
                        tmptext = columnmax[0].GetData(i + 1).ToString();
                        tmptext = (tmptext == Mtblib.Tools.MtbTools.MISSINGVALUE.ToString()) ? "*" : tmptext;
                        tttt = (tttt == "") ? (tmptext) : (tttt + "\\r" + tmptext);
                    }
                    if (CountVisible)
                    {
                        tmptext = columncount[0].GetData(i + 1).ToString();
                        tmptext = (tmptext == Mtblib.Tools.MtbTools.MISSINGVALUE.ToString()) ? "*" : tmptext;
                        tttt = (tttt == "") ? (tmptext) : (tttt + "\\r" + tmptext);
                    }
                    textBox.Text = string.Format("\"{0}\"", tttt);
                    textBox.Unit = 0;
                    //textBox.SetTextSize(TextSize);
                    textBox.SetBoxposition(xmindataregion + xgroupunit * (i + 1 / 2),
                        xmindataregion + xgroupunit * (i + 2), 0, 0.08 + countstats * 0.03);
                    cmnd.AppendLine(textBox.GetCommand());
                }
                
                #region closed raw successful version of textbox
                //for (int i = 0; i < columnname[0].RowCount; i++)
                //{
                //    //columnname[0].GetData(i + 1); //group 1 name 沒用
                //    //if (i == 0) break;
                //    if (columnname[0].GetData(i + 1) == Mtblib.Tools.MtbTools.MISSINGVALUE.ToString())
                //    {
                //        Console.WriteLine("{0}", columnname[0].GetData(i + 1));
                //        break;
                //    }
                //    tb = new Mtblib.Graph.Component.Annotation.Textbox();
                //    textBox = new MtbGraph.Component.Annotation.Adapter_Textbox(tb);
                //    textBox.SetCoordinate(xmindataregion + xgroupunit * ( i + 1/2), 0.155); // 用group總個數算x

                //    string tttt = "";
                //    for (int j = 0; j < colummean[0].RowCount; j++)
                //    {
                //        tttt =  
                //            (j == 0)?  (colummean[0].GetData(j + 1).ToString()):(tttt + "\\r" + colummean[0].GetData(j + 1).ToString());
                //        //if (j!=0) tttt = tttt + "\r" + colummean[0].GetData(j + 1).ToString();
                //    }
                //    textBox.Text = string.Format("\"{0}\"", tttt);
                //    //textBox.Text = "\"Mean\rQ1\"";
                //    textBox.Unit = 0;
                //    //double xminadd = 0.5 / columnname[0].RowCount * (i + 1);
                //    textBox.SetBoxposition(xmindataregion + xgroupunit * (i + 1 / 2), xmindataregion + (xmaxdataregion - xmindataregion)/ columnname[0].RowCount * (i + 3 / 2), 0.155, 0.25);
                //    cmnd.AppendLine(textBox.GetCommand());
                //}
                #endregion
            }
            cmnd.Append(_boxplot.GetAnnotationCommand());
            
            cmnd.AppendLine(".");
            #endregion

            return cmnd.ToString();
        }

        /// <summary>
        /// 指定各個統計量是否顯示.
        /// </summary>
        public bool IfStatVisible { get; set; }
        public bool MeanVisible { get; set; }
        public bool Q1Visible { get; set; }
        public bool Q3Visible { get; set; }
        public bool MedianVisible { get; set; }
        public bool MaxVisible { get; set; }
        public bool MinVisible { get; set; }
        public bool CountVisible { get; set; }

        /// <summary>
        /// 指定或取得圖形儲存路徑(位置+檔名+副檔名)，副檔名可以是 JPG, JPEG, MGF.
        /// </summary>
        public string GSave { set; get; }

        protected int _textsize;
        protected int TextSize
        {
            get { return _textsize; }
            set { _textsize = value; }
        }

        public void SetTextSize(dynamic size)
        {
            _textsize = size;
        }

        protected double _division;

        protected virtual void DefaultSetting()
        {
            IfStatVisible = true;
            MeanVisible = true;
            Q1Visible = true;
            Q3Visible = true;
            MedianVisible = true;
            MaxVisible = true;
            MinVisible = true;
            CountVisible = true;
            TextSize = 10;
            //_xscale = new Mtblib.Graph.Component.Scale.CateScale(Mtblib.Graph.Component.ScaleDirection.X_Axis);
            //_xscale.Label.Visible = false;
            //_xscale.Ticks.Angle = 0;
            //_title = new Mtblib.Graph.Component.Title();


            //_boxplot.CMean.Visible = true;
            //_boxplot.CMean.Size = 1.5;
            //_boxplot.Mean.Visible = true;
            //_boxplot.Mean.Type = 6; //solid circle
            //_boxplot.Mean.Color = 64; // Indigo
            //_boxplot.Mean.Size = 1.5;
            //_boxplot.Individual.Visible = true;
            //_boxplot.Individual.Color = 20; // Medium Gray

            //_boxplot.IQRBox.Visible = false;
            //_boxplot.Whisker.Visible = false;
            //_boxplot.RBox.Visible = false;
            //_boxplot.Outlier.Visible = false;
            //_boxplot.Title.Visible = false;
            //_boxplot.FigureRegion.AutoSize = true;
            ////_boxplot.FigureRegion.SetCoordinate(0, 1, Division, 1);
            //_boxplot.DataRegion.AutoSize = false;

            //_division = 0.455;
            //_boxplot.DataRegion.AutoSize = false;
            //_boxplot.DataRegion.SetCoordinate(0.1065, 0.9533, _division, 0.8814);

        }

        public void SetMtbEnvironment(Mtb.Project proj, Mtb.Worksheet ws)
        {
            _proj = proj;
            _ws = ws;
            _boxplot = new Mtblib.Graph.CategoricalChart.BoxPlot(_proj, _ws);
            //_xscale = new Mtblib.Graph.Component.Scale.CateScale(Mtblib.Graph.Component.ScaleDirection.X_Axis);
            //_boxplot = new Mtblib.Graph.CategoricalChart.BoxPlot(_proj, _ws);
            //_datlabOptAtChart = new DatalabOption();
            //_datlabOptAtBoxPlot = new DatalabOption();
            //_datlabOptAtBoxPlotIndiv = new DatalabOption();
            //SetDefault();
            _boxplot.SetDefault();
            DefaultSetting();
        } 


        //public bool Whisker { get => throw new NotImplementedException(); set => throw new NotImplementedException(); }

        protected Mtb.Column[] _groupVariables = null;
        public virtual dynamic GroupingVariables
        {
            set
            {
                if (value == null)
                {
                    _groupVariables = null;
                    return;
                }
                Mtb.Column[] _cols = Mtblib.Tools.MtbTools.GetMatchColumns(value, _ws);
                if (_cols.Length > 3)
                    throw new ArgumentException(
                        string.Format("輸入欄位數 {0}，不是合法的數量。HLBarLinePlot 用於分群的欄位數量最多為 3。", _cols.Length));
                _groupVariables = _cols;
            }
            get
            {
                return _groupVariables;
            }
        }

        /// <summary>
        /// 設定要分群的資料欄，合法的輸入為單一(string)或多個(string[], 最多3個)欄位
        /// </summary>
        /// <param name="var"></param>
        public void SetGroupingBy(dynamic var)
        {
            GroupingVariables = var;
        }
        /// <summary>
        /// 設定變數的資料欄
        /// </summary>
        /// <param name="var"></param>
        public void SetVariables(dynamic var)
        {
            Variables = var;
        }

        protected Mtb.Column[] _var = null;
        public virtual dynamic Variables
        {
            set
            {
                if (value == null)
                {
                    _var = null;
                    return;
                }
                Mtb.Column[] _cols = Mtblib.Tools.MtbTools.GetMatchColumns(value, _ws);
                if (_cols.Length > 1)
                    throw new ArgumentException(
                        string.Format("輸入欄位數 {0}，不是合法的數量。CBoxplot只能繪製一個欄位。", _cols.Length));
                _var = _cols;
            }
            get
            {
                return _var;
            }
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
                _boxplot.Dispose();
            }
            // Free your own state (unmanaged objects).
            // Set large fields to null.
            Variables = null;
            GroupingVariables = null;
            _proj = null;
            _ws = null;
            GC.Collect();

        }
        ~CBoxplot()
        {
            Dispose(false);
        }

        #region close all items
        //// protected Mtblib.Graph.Component.Scale.CateScale _xscale;
        ///// <summary>
        ///// 取得圖形中的 X 軸元件
        ///// </summary>
        //public Component.Scale.ICateScale XScale
        //{
        //    get { return new Component.Scale.Adapter_CateScale(_boxplot.XScale); }
        //}

        ///// <summary>
        ///// 取得 Boxplot 的 Y 軸元件
        ///// </summary>
        //public Component.Scale.IContScale YScale
        //{
        //    get
        //    {
        //        return new Component.Scale.Adapter_ContScale(_boxplot.YScale);
        //    }
        //}

        ///// <summary>
        ///// 取得 Boxplot Datalab 的物件
        ///// </summary>
        //public Component.IDatlab Datlab
        //{
        //    get
        //    {
        //        return new Component.Adapter_DatLab(_boxplot.MeanDatlab);
        //    }
        //}

        //public Component.Region.IRegion DataRegion
        //{
        //    get
        //    {
        //        return new Component.Region.Adapter_Region(_boxplot.DataRegion);
        //    }
        //}

        //public Component.DataView.IDataView Mean
        //{
        //    get
        //    {
        //        return new Component.DataView.Adapter_DataView(_boxplot.Mean);
        //    }
        //}

        //public Component.DataView.IDataView CMean
        //{
        //    get
        //    {
        //        return new Component.DataView.Adapter_DataView(_boxplot.CMean);
        //    }
        //}

        ////public Component.DataView.IBox RBox
        ////{
        ////    get
        ////    {
        ////        return new Component.DataView.Adapter_Box(_boxplot.RBox);
        ////    }
        ////}

        //public Component.DataView.IBox IQRBox
        //{
        //    get
        //    {
        //        return new Component.DataView.Adapter_Box(_boxplot.IQRBox);
        //    }
        //}

        ///// <summary>
        ///// 取得 Boxplot Chart Individual Datalab 的物件
        ///// </summary>
        //public Component.IDatlab DatlabAtBoxPlotIndiv
        //{
        //    get
        //    {
        //        return new Component.Adapter_DatLab(_boxplot.IndivDatlab);
        //    }
        //}

        //public Component.DataView.IDataView Individual
        //{
        //    get
        //    {
        //        return new Component.DataView.Adapter_DataView(_boxplot.Individual);
        //    }
        //}

        //public Component.DataView.IDataView Outlier
        //{
        //    get
        //    {
        //        return new Component.DataView.Adapter_DataView(_boxplot.Outlier);
        //    }
        //}

        //protected string DefaultCommand()
        //{
        //    if (Variables == null) return "";

        //    Mtb.Column[] vars = (Mtb.Column[])Variables;
        //    Mtb.Column[] gps = null;
        //    if (GroupingVariables != null)
        //    {
        //        gps = (Mtb.Column[])GroupingVariables;
        //    }

        //    StringBuilder cmnd = new StringBuilder(); // local macro 內容
        //    if (gps != null)
        //    {
        //        cmnd.AppendFormat("Boxplot ({0})*{1};\r\n",
        //            string.Join(" &\r\n", vars.Select(x => x.SynthesizedName).ToArray()),
        //            gps[0].SynthesizedName);
        //        if (gps.Length >= 2)
        //            cmnd.AppendFormat(" Group {0};\r\n",
        //            string.Join(" &\r\n", gps.Select((x, i) => new { colId = x.SynthesizedName, index = i }).
        //            Where(x => x.index > 0).Select(x => x.colId).ToArray()));
        //    }
        //    else
        //    {
        //        cmnd.AppendFormat("Boxplot {0};\r\n",
        //            string.Join(" &\r\n", vars.Select(x => x.SynthesizedName).ToArray()));
        //    }

        //    if (GSave != null)
        //    {
        //        _boxplot.GraphPath = GSave;
        //        cmnd.Append(_boxplot.GetOptionCommand());
        //        _boxplot.GraphPath = null;
        //    }

        //    cmnd.Append(YScale.GetCommand());
        //    cmnd.Append(XScale.GetCommand());

        //    cmnd.Append(Mean.GetCommand());
        //    cmnd.Append(CMean.GetCommand());
        //    cmnd.Append(RBox.GetCommand());
        //    cmnd.Append(IQRBox.GetCommand());
        //    //cmnd.Append(Whisker.GetCommand());
        //    cmnd.Append(Outlier.GetCommand());
        //    cmnd.Append(Individual.GetCommand());
        //    cmnd.Append(MeanDatlab.GetCommand());
        //    cmnd.Append(IndivDatlab.GetCommand());

        //    cmnd.Append(Panel.GetCommand());

        //    cmnd.Append(GetAnnotationCommand());
        //    cmnd.Append(GetRegionCommand());

        //    return cmnd.ToString() + ".";
        //}
        #endregion
    }
}
