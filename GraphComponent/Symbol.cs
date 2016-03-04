using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    [ClassInterface(ClassInterfaceType.None)]
    public class Symbol : ICOMInterop_Line, IDataView, IGroup
    {
        //private Object sType;
        //private Object sColor;
        //private Object sSize;


        public Symbol()
        {
            //sType = null;
            //sColor = null;
            //sSize = null;
        }

        public Symbol Clone()
        {
            return (Symbol)this.MemberwiseClone();
        }

        //public void SetType(ref Object type)
        //{
        //    this.sType = type;
        //}
        //public void SetColor(ref Object color)
        //{
        //    this.sColor = color;
        //}
        //public void SetSize(ref Object size)
        //{
        //    this.sSize = size;
        //}
        public void SetDefault()
        {
            //sType = null;
            //sColor = null;
            //sSize = null;
            _type = null;
            _color = null;
            _size = null;
            _groupBy = null;

        }
        /// <summary>
        /// 取得 Symbol 的 Color
        /// </summary>
        /// <returns></returns>
        public int[] GetColor()
        {
            //return this.sColor;
            return _color;
        }
        /// <summary>
        /// 取得 Symbol 的 Size 
        /// </summary>
        /// <returns></returns>
        public int[] GetSize()
        {
            //return this.sSize;
            return _size;
        }
        /// <summary>
        /// 取得 Symbol 的 Type
        /// </summary>
        /// <returns>int[]</returns>
        public int[] GetTypes()
        {
            //return this.sType;
            return _type;
        }

        private String _groupBy;
        public void SetGroupBy(String byColumn)
        {
            _groupBy = byColumn;
        }


        private StringBuilder cmnd = new StringBuilder();
        public virtual String GetCommand()
        {
            cmnd.Clear();
            cmnd.Append(" SYMBOL" + (_groupBy == null ? ";" : " " + _groupBy + ";") + Environment.NewLine);
            if (_type != null)
            {
                cmnd.AppendLine(string.Format("  Type {0};", string.Join(" ", _type)));
            }
            if (_color != null)
            {
                cmnd.AppendLine(string.Format("  Color {0};", string.Join(" ", _color)));
            }
            if (_size != null)
            {
                cmnd.AppendLine(string.Format("  Size {0};", string.Join(" ", _size)));
            }

            //Type t;
            //if (this.sType != null)
            //{
            //    t = this.sType.GetType();
            //    if (t.IsArray)
            //    {                    
            //        try
            //        {
            //            IEnumerable em = this.sType as IEnumerable;
            //            cmnd.Append("  TYPE ");
            //            foreach (object o in em)
            //            {
            //                cmnd.Append(o.ToString() + " ");
            //            }
            //            cmnd.Append(";" + Environment.NewLine);

            //        }
            //        catch 
            //        {
            //            throw new ArgumentException("Invalid Type value of Symbol");
            //        }
            //    }
            //    else
            //    {
            //        Console.WriteLine("It is not array, {0}", this.sType.ToString());
            //        cmnd.Append("  TYPE " + this.sType.ToString() + ";" + Environment.NewLine);
            //    }

            //}
            //if (this.sColor != null)
            //{
            //    t = this.sColor.GetType();
            //    Console.WriteLine(t.ToString());
            //    if (t.IsArray)
            //    {
            //        Console.WriteLine("It is array");
            //        try
            //        {
            //            IEnumerable em = this.sColor as IEnumerable;
            //            StringBuilder tmpSb = new StringBuilder();
            //            cmnd.Append("  COLOR ");
            //            foreach (object o in em)
            //            {
            //                cmnd.Append(o.ToString() + " ");
            //            }
            //            cmnd.Append(";" + Environment.NewLine);
            //        }
            //        catch
            //        {
            //            throw new ArgumentException("Invalid Color value of Symbol");
            //        }
            //    }
            //    else
            //    {
            //        cmnd.Append("  COLOR " + this.sColor.ToString() + ";" + Environment.NewLine);
            //    }

            //}
            //if (this.sSize != null)
            //{
            //    t = this.sSize.GetType();
            //    Console.WriteLine(t.ToString());
            //    if (t.IsArray)
            //    {
            //        Console.WriteLine("It is array");
            //        try
            //        {
            //            IEnumerable em = this.sColor as IEnumerable;
            //            StringBuilder tmpSb = new StringBuilder();
            //            cmnd.Append("  SIZE ");
            //            foreach (object o in em)
            //            {
            //                cmnd.Append(o.ToString() + " ");
            //            }
            //            cmnd.Append(";" + Environment.NewLine);
            //        }
            //        catch
            //        {
            //            throw new ArgumentException("Invalid Size value of Symbol");
            //        }
            //    }
            //    else
            //    {
            //        cmnd.Append("  SIZE " + this.sSize.ToString() + ";" + Environment.NewLine);
            //    }
            //}

            return cmnd.ToString();
        }


        private int[] _type = null;
        /// <summary>
        /// 設定 Symbol 的類型，輸入整數陣列
        /// </summary>
        /// <param name="type"></param>
        public void SetType(dynamic type)
        {
            _type = ConvertInputToArray(type);
        }

        private int[] _color = null;
        public void SetColor(dynamic color)
        {
            _color = ConvertInputToArray(color);
        }

        private int[] _size = null;
        public void SetSize(dynamic size)
        {
            _size = ConvertInputToArray(size);
        }

        private int[] ConvertInputToArray(dynamic d)
        {
            int[] result = null;
            Type t = d.GetType();
            if (t.IsArray)
            {
                try
                {
                    IEnumerable em = d as IEnumerable;
                    result = em.Cast<object>().Select(x => int.Parse(x.ToString())).ToArray();
                    
                }
                catch (Exception ex)
                {
                    throw new ArgumentException("Invalid Type value of Symbol");
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
                    throw new ArgumentException("Invalid Type value of Symbol");
                }
            }
            return result;
        }

    }

}
