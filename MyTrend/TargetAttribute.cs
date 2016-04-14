using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.MyTrend
{
    [System.Runtime.InteropServices.ClassInterface(System.Runtime.InteropServices.ClassInterfaceType.None)]
    public class TargetAttribute : ICOMInterop_TargetAttribute, GraphComponent.IDataView
    {
        public TargetAttribute()
        {

        }

        private bool _showTargetNotation = true;
        /// <summary>
        /// 用於決定 Target 的註解是否要顯示，預設值為 True
        /// </summary>
        public bool ShowNotation
        {
            set
            {
                _showTargetNotation = value;
            }
            get
            {
                return _showTargetNotation;
            }
        }

        private int[] _color = null;
        /// <summary>
        /// 設定 Target line 的顏色屬性，int[]
        /// </summary>
        /// <param name="color"></param>
        public void SetColor(dynamic color)
        {
            _color = ConvertInputToArray(color);

        }

        private int[] _type = null;
        /// <summary>
        /// 設定 Target line 的線段類型，int[]
        /// </summary>
        /// <param name="linetype"></param>
        public void SetType(dynamic linetype)
        {
            _type = ConvertInputToArray(linetype);
        }

        private int[] _size = null;
        /// <summary>
        /// 設定 Target line 的線段粗細，int[]
        /// </summary>
        /// <param name="size"></param>
        public void SetSize(dynamic size)
        {
            _size = ConvertInputToArray(size);
        }

        private int[] _fontSize = null;
        public void SetNotationSize(dynamic fontSize)
        {
            _fontSize = ConvertInputToArray(fontSize);
        }


        /// <summary>
        /// 取得 Target 的 Type
        /// </summary>
        /// <returns></returns>
        public int[] GetTypes()
        {
            return _type;
        }

        /// <summary>
        /// 取得 Target 的 Color
        /// </summary>
        /// <returns></returns>
        public int[] GetColor()
        {
            return _color;
        }

        /// <summary>
        /// 取得 Target 的 Size
        /// </summary>
        /// <returns></returns>
        public int[] GetSize()
        {
            return _size;
        }

        /// <summary>
        /// 取得 Target 的文字資訊的字型大小
        /// </summary>
        /// <returns></returns>
        public int[] GetNotationSize()
        {
            return _fontSize;
        }

        public string GetCommand()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// 將 Target 的屬性設定回復初始值
        /// </summary>
        public void SetDefault()
        {
            _color = null;
            _type = null;
            _size = null;
            _fontSize = null;
            _showTargetNotation = true;

        }

        public TargetAttribute Clone()
        {
            return (TargetAttribute)this.MemberwiseClone();
        }

        private int[] ConvertInputToArray(dynamic d)
        {
            int[] result = null;
            Type t = d.GetType();
            if (t.IsArray)
            {
                try
                {
                    System.Collections.IEnumerable em = d as System.Collections.IEnumerable;
                    result = em.Cast<object>().Select(x => int.Parse(x.ToString())).ToArray();

                }
                catch (Exception ex)
                {
                    throw new ArgumentException("Invalid Type value of input variable");
                }
            }
            else
            {
                try
                {
                    result = new int[] { int.Parse(d.ToString()) };
                }
                catch (Exception ex)
                {
                    throw new ArgumentException("Invalid Type value of input variable.");
                }
            }
            return result;
        }

        
    }
}
