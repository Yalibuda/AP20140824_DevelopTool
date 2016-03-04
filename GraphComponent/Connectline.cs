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
    public class Connectline : ICOMInterop_Line,IDataView, IGroup
    {
        //private Object cType;
        //private Object cColor;
        //private Object cSize;
        public Connectline()
        {
            //cType = null;
            //cColor = null;
            //cSize = null;
        }

        public Connectline Clone()
        {
            return (Connectline)this.MemberwiseClone();
        }

        //public void SetType(ref Object type)
        //{
        //    this.cType = type;
        //}
        //public void SetColor(ref Object color)
        //{
        //    this.cColor = color;
        //}
        //public void SetSize(ref Object size)
        //{
        //    this.cSize = size;
        //}
        public void SetDefault()
        {
            //cType = null;
            //cColor = null;
            //cSize = null;
            _color = null;
            _type = null;
            _size = null;
            _groupBy = null;
        }

        public int[] GetTypes()
        {
            return _type;
        }

        public int[] GetColor()
        {
            return _color;
        }

        public int[] GetSize()
        {
            return _size;
        }

        private StringBuilder cmnd = new StringBuilder();
        public virtual String GetCommand()
        {
            cmnd.Clear();
            cmnd.AppendLine(" CONN" + (this._groupBy == null ? ";" : " " + this._groupBy + ";"));
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
            //if (this.cType != null)
            //{
            //    t = this.cType.GetType();
            //    if (t.IsArray)
            //    {
            //        Console.WriteLine("It is array");
            //        try
            //        {
            //            IEnumerable em = this.cType as IEnumerable;
            //            StringBuilder tmpSb = new StringBuilder();
            //            cmnd.Append("  TYPE ");
            //            foreach (object o in em)
            //            {
            //                cmnd.Append(o.ToString() + " ");
            //            }
            //            cmnd.AppendLine(";");

            //        }
            //        catch (Exception e)
            //        {
            //            throw new Exception(e.Message);
            //        }
            //    }
            //    else
            //    {
            //        Console.WriteLine("It is not array, {0}", this.cType.ToString());
            //        cmnd.Append("  TYPE " + this.cType.ToString() + ";" + Environment.NewLine);
            //    }

            //}
            //if (this.cColor != null)
            //{
            //    t = this.cColor.GetType();
            //    Console.WriteLine(t.ToString());
            //    if (t.IsArray)
            //    {
            //        Console.WriteLine("It is array");
            //        try
            //        {
            //            IEnumerable em = this.cColor as IEnumerable;
            //            StringBuilder tmpSb = new StringBuilder();
            //            cmnd.Append("  COLOR ");
            //            foreach (object o in em)
            //            {
            //                cmnd.Append(o.ToString() + " ");
            //            }
            //            cmnd.Append(";" + Environment.NewLine);
            //        }
            //        catch (Exception e)
            //        {
            //            throw new Exception(e.Message);
            //        }
            //    }
            //    else
            //    {
            //        cmnd.Append("  COLOR " + this.cColor.ToString() + ";" + Environment.NewLine);
            //    }

            //}
            //if (this.cSize != null)
            //{
            //    t = this.cSize.GetType();
            //    Console.WriteLine(t.ToString());
            //    if (t.IsArray)
            //    {
            //        Console.WriteLine("It is array");
            //        try
            //        {
            //            IEnumerable em = this.cSize as IEnumerable;
            //            StringBuilder tmpSb = new StringBuilder();
            //            cmnd.Append("  SIZE ");
            //            foreach (object o in em)
            //            {
            //                cmnd.Append(o.ToString() + " ");
            //            }
            //            cmnd.Append(";" + Environment.NewLine);
            //        }
            //        catch (Exception e)
            //        {
            //            throw new Exception(e.Message);
            //        }
            //    }
            //    else
            //    {
            //        cmnd.Append("  SIZE " + this.cSize.ToString() + ";" + Environment.NewLine);
            //    }
            //}

            return cmnd.ToString();
        }


        private String _groupBy;
        public void SetGroupBy(string byColumn)
        {
            this._groupBy = byColumn;
        }


        private int[] _type = null;
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
