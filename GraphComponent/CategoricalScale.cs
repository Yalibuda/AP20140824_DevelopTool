using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class CategoricalScale : ContinuousScale, ICOMInterop_Scale_Categorical
    {

        public CategoricalScale(ScaleType scale_axis)
            : base(scale_axis)
        {
            this.scale_axis = scale_axis;
        }

        public override IScale Clone()
        {
            CategoricalScale catescale = new CategoricalScale(this.scale_axis);
            catescale.Min = this.Min.Clone();
            catescale.Max = this.Max.Clone();
            catescale.Tick = this.Tick.Clone();
            catescale.AxLab = this.AxLab.Clone();
            catescale.Reference = this.Reference.Clone();
            return catescale;
        }

        public override String GetCommand()
        {
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
            StringBuilder cmnd = new StringBuilder();
            StringBuilder sb = new StringBuilder();
            if (k == 0)
            {
                throw new ArgumentException("Invalid Scale type, it should be X-axis, Y-axis");
                return null;
            }
            else
            {
                if (this.Min.Value < 1.23456E+30)
                {
                    sb.AppendLine("  MIN " + this.Min.Value + ";");
                }
                if (this.Max.Value < 1.23456E+30)
                {
                    sb.AppendLine("  MAX " + this.Max.Value + ";");
                }
                switch (((ScaleTick)Tick).TickAttr)
                {
                    case TickAttribute.NumberOfTicks:
                        if (this.datacount == -1)
                        {
                            sb.Append("");
                         }
                        sb.AppendLine("  TINC " + Math.Max(Math.Ceiling((double)this.datacount / (double)Tick.GetNumberOfMajorTick()), 1) + ";");
                        break;
                    case TickAttribute.ByIncrement:
                        sb.AppendLine("  TINC " + Tick.GetIncrement() + ";" );
                        break;
                    case TickAttribute.ShowAllTickAsPossible:
                        sb.AppendLine("  TINC " + (datacount <= 54 ? 1 : Math.Ceiling((double)this.datacount / (double)54)) + ";");
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
                if (this.scale_axis == ScaleType.Secondary_Y_axis & this.varStr != String.Empty)
                {
                    cmnd.AppendLine("  SECS " + this.varStr + ";");
                }

                /*
                 * 非 Secondary scale 才需要 call 裡面的 Secondary 物件建立 Secondary 
                 * 的 Scale command
                 * 
                 */
                if (this.varStr != String.Empty && this.scale_axis != ScaleType.Secondary_Y_axis)
                {
                    try
                    {
                        cmnd.Append(SecsScale.GetCommand());
                    }
                    catch
                    {
                        cmnd.AppendLine("#This scale do not have secondary scale");
                    }
                    
                }

                /*
                 * 要留意，如果生成 Secondary 的 Axlabel 若無設定 Secondary 變數會出現錯誤。
                 * 
                 */
                if (this.scale_axis == ScaleType.Secondary_Y_axis)
                {
                    if (this.varStr != String.Empty) cmnd.Append(AxLab.GetCommand());
                }
                else
                {
                    cmnd.Append(AxLab.GetCommand());
                    cmnd.Append(Reference.GetCommand());
                }

                return cmnd.ToString();
            }

        }


        private long datacount = -1;
        protected override void CalculateTickValue(List<String> varCols, Mtb.Worksheet ws, Mtb.Project proj)
        {
            long row;
            foreach (String str in varCols)
            {
                row = ws.Columns.Item(str).RowCount;
                if (row > datacount) datacount = row;
            }

        }
    }
}
