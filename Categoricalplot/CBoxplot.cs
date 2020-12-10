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
        //Composite Boxplot
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
            DefaultSetting();
            string cmnd = GetCommand();
            string fpath = Mtblib.Tools.MtbTools.BuildTemporaryMacro("mycode.mtb", cmnd.ToString());
            _proj.ExecuteCommand(string.Format("exec \"{0}\" 1", fpath));
        }

        protected virtual string GetCommand()
        {
            Mtb.Column[] vars;
            if (Variables == null)
            {
                throw new ArgumentNullException("Variable 不可為空");
            }
            else
            {
                vars = Variables;
            }

            Mtb.Column[] gps = GroupingVariables;

            #region compute stats

            #endregion

            #region Create Boxplot
            StringBuilder cmnd = new StringBuilder();
            if (gps != null)
            {
                cmnd.AppendFormat("Boxplot ({0})*{1};\r\n",
                    string.Join(" &\r\n", vars.Select(x => x.SynthesizedName).ToArray()),
                    gps[0].SynthesizedName);
                if (gps.Length >= 2)
                    cmnd.AppendFormat(" Group {0};\r\n",
                    string.Join(" &\r\n", gps.Select((x, i) => new { colId = x.SynthesizedName, index = i }).
                    Where(x => x.index > 0).Select(x => x.colId).ToArray()));
            }
            else
            {
                cmnd.AppendFormat("Boxplot {0};\r\n",
                    string.Join(" &\r\n", vars.Select(x => x.SynthesizedName).ToArray()));
            }

            cmnd.Append(_boxplot.GetOptionCommand());

            cmnd.Append(_boxplot.YScale.GetCommand());
            cmnd.Append(_boxplot.XScale.GetCommand());

            cmnd.Append(_boxplot.Mean.GetCommand());
            cmnd.Append(_boxplot.CMean.GetCommand());
            cmnd.Append(_boxplot.RBox.GetCommand());
            cmnd.Append(_boxplot.IQRBox.GetCommand());
            cmnd.Append(_boxplot.Whisker.GetCommand());
            cmnd.Append(_boxplot.Outlier.GetCommand());
            cmnd.Append(_boxplot.Individual.GetCommand());
            cmnd.Append(_boxplot.MeanDatlab.GetCommand());
            cmnd.Append(_boxplot.IndivDatlab.GetCommand());

            cmnd.Append(_boxplot.Panel.GetCommand());

            #region textbox auto computation
            Mtblib.Graph.Component.Annotation.Textbox tb = new Mtblib.Graph.Component.Annotation.Textbox();
            tb.SetCoordinate(0, -5.655);
            tb.Text = "\"aaaaaa\raaa\raa\ra\"";
            tb.Unit = 1;
            tb.Angle = 0;
            // placement 跟 offset很難用, 自己寫textbox語法彌補沒有box位置/大小調整
            //tb.SetCoordinate(0, 5, -4, -7);
            _boxplot.ATextLst.Add(tb);
            #endregion
            cmnd.Append(_boxplot.GetAnnotationCommand());

            _division = 0.455;
            _boxplot.DataRegion.AutoSize = false;
            _boxplot.DataRegion.SetCoordinate(0.1065, 0.9533, _division, 0.8814);
            cmnd.Append(_boxplot.GetRegionCommand());

            

            cmnd.AppendLine(".");
            #endregion

            
            return cmnd.ToString();
        }



        /// <summary>
        /// 指定或取得圖形儲存路徑(位置+檔名+副檔名)，副檔名可以是 JPG, JPEG, MGF.
        /// </summary>
        public string GSave { set; get; }

        protected double _division;

        protected virtual void DefaultSetting()
        {
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
            //PanelBy = null;
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
