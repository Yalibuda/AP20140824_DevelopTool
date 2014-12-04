using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class ContinuousScale : Graphcomp, ICOMInterop_Scale_Continuous, IScale
    {
        /*
         * Initialize 
         */
        protected ScaleType scale_axis;
        public ContinuousScale(ScaleType scale_axis)
        {
            this.scale_axis = scale_axis;
            this.Min = new ScaleBoundary();
            this.Max = new ScaleBoundary();
            this.Tick = new ScaleTick();
            this.AxLab = new AxLabel(this.scale_axis);
            this.Reference = new Reference(this.scale_axis);
            /*
             * MyGSacle 這是一個很麻煩的東西，其目的主要是算出座標軸的
             * 邊界、Tick 標籤等資訊。因此，它需要參考 proj, ws 運作。
             * 產生的麻煩...Scale 需要參考
             */

            this.GScale = new MyGScale();
            //this.secsInfo = "";
            /*
             * 建立 Secondary Scale
             */
            if (scale_axis == ScaleType.Y_axis)
            {
                SecsScale = new ContinuousScale(ScaleType.Secondary_Y_axis);
            }
        }

        public virtual IScale Clone()
        {
            /*
             * 此方法使用時要留意，回傳並非 ContinuouScale 而是 IScale，所以
             * 若有其他用途記得要降轉。
             */
            ContinuousScale contscale = new ContinuousScale(this.scale_axis);
            contscale.Min = this.Min.Clone();
            contscale.Max = this.Max.Clone();
            contscale.Tick = this.Tick.Clone();
            contscale.AxLab = this.AxLab.Clone();
            contscale.Reference = this.Reference.Clone();
            contscale.GScale = this.GScale.Clone();
            return contscale;
        }

        public ScaleBoundary Min { set; get; }
        public ScaleBoundary Max { set; get; }
        public ScaleTick Tick { set; get; }
        public AxLabel AxLab { set; get; }
        public Reference Reference { set; get; }
        public ContinuousScale SecsScale { private set; get; }
        public MyGScale GScale { private set; get; }
        public bool ShowHighSide { set; get; }

        public virtual String GetCommand()
        {
            StringBuilder cmnd = new StringBuilder();
            int k = 0;
            switch (scale_axis)
            {
                case ScaleType.X_axis:
                    k = 1;
                    break;
                case ScaleType.Y_axis:
                case ScaleType.Secondary_Y_axis:
                    k = 2;
                    break;
            }
            StringBuilder sb = new StringBuilder();
            if (k == 0)
            {
                throw new ArgumentException("Invalid Scale type, it should be X-axis, Y-axis");
                return null;
            }
            else
            {
                if (this.ShowHighSide)
                {
                    sb.AppendLine("  LDISP 1 0 0 0;" + Environment.NewLine + "  HDISP 1 1 1 0;");
                }

                //cmnd.Append(" SCALE " + k + ";" + Environment.NewLine);
                if (this.Min.Value < 1.23456E+30)
                {
                    sb.Append("  MIN " + this.Min.Value + ";" + Environment.NewLine);
                }
                if (this.Max.Value < 1.23456E+30)
                {
                    sb.Append("  MAX " + this.Max.Value + ";" + Environment.NewLine);
                }
                switch (((ScaleTick)Tick).TickAttr)
                {
                    case TickAttribute.NumberOfTicks:
                        sb.Append("  NMAJ " + Tick.GetNumberOfMajorTick() + ";" + Environment.NewLine);
                        break;
                    case TickAttribute.ByIncrement:

                        if ((this.GScale.TMIN == 1.23456E+30 & Min.Value == 1.23456E+30) ||
                            (this.GScale.TMAX == 1.23456E+30 & Max.Value == 1.23456E+30))
                        {
                            sb.AppendLine("  #未輸入變數或Scale邊界，無法使用 interval 方法");
                        }
                        else
                        {
                            sb.AppendLine("  TICK " + (Min.Value == 1.23456E+30 ? this.GScale.TMIN : Min.Value)
                                + ":" + (Max.Value == 1.23456E+30 ? this.GScale.TMAX : Max.Value) +
                                "/" + Tick.GetIncrement() + ";");
                        }
                        break;
                }

                if (Tick.TickAngle < 1.23456E+30)
                {
                    sb.AppendLine("  ANGLE " + Tick.TickAngle + ";");
                }

                if (Tick.FontSize != 8)
                {
                    sb.AppendLine("  PSIZE " + Tick.FontSize + ";");
                }

                if (sb.ToString() != String.Empty)
                {
                    cmnd.AppendLine(" SCALE " + k + ";");
                    cmnd.Append(sb.ToString());
                }
                /*
                 * 如果是 Secondary，切記 secs 資訊是由 primary scale 傳給 secondary scale                 
                 * 
                 */
                if (this.scale_axis == ScaleType.Secondary_Y_axis & varStr != String.Empty)
                {
                    if (sb.ToString() == String.Empty)
                    {
                        cmnd.AppendLine(" SCALE " + k + ";");
                    }
                    cmnd.AppendLine("  SECS " + varStr + ";");
                }

                /*
                 * 非 Secondary scale 才需要 call 裡面的 Secondary 物件建立 Secondary 
                 * 的 Scale command
                 * 
                 */
                if (varStr != String.Empty && this.scale_axis != ScaleType.Secondary_Y_axis)
                {
                    cmnd.Append(SecsScale.GetCommand());
                }

                /*
                 * 要留意，如果生成 Secondary 的 Axlabel 若無設定 Secondary 變數會出現錯誤。
                 * 
                 */
                if (this.scale_axis == ScaleType.Secondary_Y_axis)
                {
                    if (this.varStr != String.Empty)
                    {
                        cmnd.Append(AxLab.GetCommand());
                        cmnd.Append(Reference.GetCommand());
                    }
                }
                else
                {
                    cmnd.Append(AxLab.GetCommand());
                    cmnd.Append(Reference.GetCommand());
                }
                return cmnd.ToString();
            }

        }

        /*
         * 因為 Scale 的 tick 需要資料處理，所以要Worksheet 資訊
         * 因為採用 Minitab 內建的 Gscale 所以
         * 
         */
        protected List<String> varCols;
        protected String varStr = String.Empty;
        public void SetScaleVariable(ref object varCols, Mtb.Worksheet ws, Mtb.Project proj = null)
        {

            Type t = varCols.GetType();
            List<String> cols = new List<String>();
            DialogAppraiser da = new DialogAppraiser();
            if (varCols == null) return;
            if (t.IsArray)
            {
                try
                {
                    IEnumerable enumerable = varCols as IEnumerable;
                    foreach (object o in enumerable)
                    {
                        cols.Add(o.ToString());
                    }
                    this.varCols = da.GetMtbCols(cols, ws);
                    varStr = String.Join(" ", cols.ToArray());
                }
                catch
                {
                    throw new ArgumentException("Invalid input of scale variables");
                    return;
                }

            }
            else
            {
                varStr = varCols.ToString();
                List<String> list = new List<String>();
                list = da.GetMtbColInfo(varCols.ToString());
                list = da.GetMtbCols(list, ws);
                this.varCols = list;
            }
            CalculateTickValue(this.varCols, ws, proj);
        }

        /*
         * 利用此方法配合已輸入好的 varCols (List<String>) 抓取資料計算不同 Tick 可能會用到的資訊
         * Continuous 是  start 和 end，Categorical 則是 increment
         * 
         * 2014/11/29
         * 要特別留意此方法的缺陷，第一次畫圖使用變數 C1, C2，也在 SetScaleVariable 中設定 C1, C2，
         * 在延續相同設定下繪製 C3, C4，但是卻沒有變更 SetVariable 設定的狀況。         
         * ==> 解法: 在輸入圖形變數時，呼叫出 SetVariable 同步更新，接口端應該將 SetVaraible 隱藏         
         * 
         */
        //private double start;
        //private double end;
        //protected virtual void CalculateTickValue(List<String> varCols, Mtb.Worksheet ws)
        //{
        //    double[] datas = new double[0];
        //    double[] data;
        //    foreach (String str in varCols)
        //    {
        //        data = ws.Columns.Item(str).GetData();
        //        datas = datas.Concat(data).ToArray();
        //    }

        //    double min = datas.Min();
        //    double max = datas.Max();
        //    MyGScale gscale;
        //    gscale = new MyGScale(min, max);
        //    this.start = gscale.SMIN;
        //    this.end = gscale.SMAX;
        //}
        
        protected virtual void CalculateTickValue(List<String> varCols, Mtb.Worksheet ws, Mtb.Project proj)
        {
            if (ws == null || proj == null) return;

            double[] datas = new double[0];
            double[] data;
            foreach (String str in varCols)
            {
                data = ws.Columns.Item(str).GetData();
                datas = datas.Concat(data).ToArray();
            }

            double min = datas.Min();
            double max = datas.Max();
            this.GScale.Run(min, max, proj, ws);
        }

    }
}
