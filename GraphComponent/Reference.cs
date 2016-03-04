using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class Reference : ICOMInterop_Refe
    {
        private enum RefComponemt
        {
            Value, Color, Type
        }
        private ScaleType scale_axis;
        public Reference(ScaleType scale_axis)
        {
            this.scale_axis = scale_axis;
            this.Side = 2;
            this.FontSize = -1;
            this.Size = -1;

        }
        /*
         * 20150129:
         * 新增 haveValues 方法，讓舊版 Bar-line plot 可以
         * 增加 Reference line
         * ...未來可以考慮刪除
         */

        public bool haveValues()
        {
            if (this.value != null)
            {
                return true;
            }
            else
            {
                return false;
            }
        }

        public long Side { set; private get; }//20150129: 新增的屬性，可以設定輔助線顯示的位置
        public Reference Clone()
        {
            Reference refe = new Reference(this.scale_axis);
            if (this.value != null)
            {
                refe.value = new List<String>();
                foreach (String s in this.value)
                    refe.value.Add(s);
            }
            if (this.color != null)
            {
                refe.color = new List<String>();
                foreach (String s in this.color)
                    refe.color.Add(s);
            }
            if (this.refType != null)
            {
                refe.refType = new List<String>();
                foreach (String s in this.refType)
                    refe.refType.Add(s);
            }

            if (this._fontSize != -1) refe.FontSize = this._fontSize;
            if (this._size != -1) refe.Size = this._size;

            return refe;
        }

        private List<String> value;
        public void SetValue(ref Object value)
        {
            Type t = value.GetType();
            List<String> list = new List<String>();
            if (t.IsArray)
            {
                try
                {
                    IEnumerable enumerable = value as IEnumerable;
                    foreach (Object o in enumerable)
                    {
                        list.Add(o.ToString());
                    }
                    this.value = list;

                }
                catch
                {
                    throw new ArgumentException("Invalid input of reference value");
                    return;
                }
            }
            else
            {

                this.value = RegExOfRefe(value.ToString());
            }

        }

        private List<String> color;
        public void SetColor(ref Object value)
        {
            Type t = value.GetType();
            List<String> list = new List<String>();
            if (t.IsArray)
            {
                try
                {
                    IEnumerable enumerable = value as IEnumerable;
                    foreach (Object o in enumerable)
                    {
                        list.Add(o.ToString());
                    }
                    this.color = list;
                }
                catch
                {
                    throw new ArgumentException("Invalid input of reference color");
                    return;
                }
            }
            else if (Type.GetTypeCode(t) == TypeCode.String)
            {
                this.color = RegExOfRefe(value.ToString(), RefComponemt.Color);
            }
        }

        private List<String> refType;
        public void SetType(ref Object value)
        {
            Type t = value.GetType();
            if (t.IsArray)
            {
                try
                {
                    IEnumerable enumerable = value as IEnumerable;
                    List<String> list = new List<string>();
                    foreach (Object o in enumerable)
                    {
                        list.Add(o.ToString());
                    }
                    this.refType = list;
                }
                catch
                {
                    throw new ArgumentException("Invalid input of reference type");
                    return;
                }
            }
            else
            {
                refType = RegExOfRefe(value.ToString(), RefComponemt.Type);
            }
        }

        public bool HideLabel { set; get; }

        private int _fontSize = -1;
        public int FontSize
        {
            set
            {
                _fontSize = value;
            }
            get
            {
                return _fontSize;
            }
        }

        private int _size = -1;
        public int Size
        {
            set
            {
                _size = value;
            }
            get
            {
                return _size;
            }
        }

        public void Clear()
        {
            value = null;
            color = null;
            refType = null;
            _fontSize = -1;
            _size = -1;

        }

        public String GetCommand()
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

            List<String> list;
            Type t;
            /*
             * Convert value object
             */
            if (this.value == null)
            {
                return String.Empty;
            }
            else
            {
                RefeStatus refestatus = CheckRefeStatus();
                switch (refestatus)
                {
                    case RefeStatus.Simple:
                        cmnd.AppendLine(" REFE " + k + " " + String.Join(" ", this.value.ToArray()) + ";");
                        if (this.color != null) cmnd.AppendLine("  COLOR " + String.Join(" ", this.color.ToArray()) + ";");
                        if (this.refType != null) cmnd.AppendLine("  TYPE " + String.Join(" ", this.refType.ToArray()) + ";");
                        if (this._size > 0) cmnd.AppendLine(string.Format("  Size {0};", this._size));
                        if (this.scale_axis == ScaleType.Secondary_Y_axis) cmnd.AppendLine("  SECS;");
                        cmnd.AppendLine("  Side " + this.Side + ";");
                        if (this.HideLabel)
                        {
                            cmnd.AppendLine("  LABEL \"\";");
                        }
                        else
                        {
                            if (this._fontSize > 0)
                                cmnd.AppendLine(string.Format("  PSize {0};", this._fontSize));
                        }
                        break;
                    case RefeStatus.Multi:
                        String[] array1 = new String[value.Count()];
                        String[] array2 = new String[value.Count()];
                        if (this.color != null)
                        {
                            if (this.color.Count() == 1)
                            {
                                for (int i = 0; i < array1.Length; i++)
                                {
                                    array1[i] = this.color[0];
                                }
                            }
                            else
                            {
                                array1 = this.color.ToArray();
                            }
                        }
                        if (this.refType != null)
                        {
                            if (this.refType.Count() == 1)
                            {
                                for (int i = 0; i < array2.Length; i++)
                                {
                                    array2[i] = this.refType[0];
                                }
                            }
                            else
                            {
                                array2 = this.refType.ToArray();
                            }
                        }

                        for (int i = 0; i < this.value.Count(); i++)
                        {
                            cmnd.AppendLine(" REFE " + k + " " + value[i] + ";");
                            if (this.color != null) cmnd.AppendLine("  COLOR " + array1[i] + ";");
                            if (this.refType != null) cmnd.AppendLine("  TYPE " + array2[i] + ";");
                            if (this._size > 0) cmnd.AppendLine(string.Format("  Size {0};", this._size));
                            if (this.scale_axis == ScaleType.Secondary_Y_axis) cmnd.AppendLine("  SECS;");
                            cmnd.AppendLine("  Side " + this.Side + ";");
                            if (this.HideLabel)
                            {
                                cmnd.AppendLine("  LABEL \"\";");
                            }
                            else
                            {
                                if (this._fontSize > 0)
                                    cmnd.AppendLine(string.Format("  PSize {0};", this._fontSize));
                            }
                        }
                        break;
                    default:
                        cmnd.AppendLine(@"#輔助線和屬性數量不符");
                        break;
                }
                return cmnd.ToString();
            }
        }

        private RefeStatus CheckRefeStatus()
        {

            if (value == null)
            {
                return RefeStatus.None;
            }
            else
            {
                int count = this.value.Count();
                int colorCount = (this.color == null ? 0 : this.color.Count());
                int typeCount = (this.refType == null ? 0 : this.refType.Count());
                if (colorCount <= 1 & typeCount <= 1)
                {
                    return RefeStatus.Simple;
                }
                else if ((colorCount > 1 && colorCount != count) || (typeCount > 1 && typeCount != count))
                {
                    return RefeStatus.None;
                }
                else
                {
                    return RefeStatus.Multi;
                }
            }
        }


        private List<String> RegExOfRefe(String inputStr, RefComponemt refcomp = RefComponemt.Value)
        {
            Regex regEx;
            bool b;
            switch (refcomp)
            {
                case RefComponemt.Color:
                case RefComponemt.Type:
                    regEx = new Regex(@"[^/:\s]+");
                    break;
                default:
                    regEx = new Regex(@"[^/:\s]+\s*:\s*[^/:\s]+\s*/\s*[^/:\s]+|[^/:\s]+\s*:\s*[^/:\s]+|[^/:\s]+");
                    break;
            }

            b = regEx.IsMatch(inputStr);
            /*
             * 使用序列表示的視為一個Value，對應一組color 和 type，
             * 對於color, type 不考慮序列語法
             */
            List<String> list = new List<string>();
            if (regEx.IsMatch(inputStr))
            {
                MatchCollection matchs = regEx.Matches(inputStr);
                foreach (Match m in matchs)
                {
                    list.Add(m.Value);
                }
            }
            return list;
        }

    }
}
