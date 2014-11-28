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
        private Object sType;
        private Object sColor;
        private Object sSize;


        public Symbol()
        {
            sType = null;
            sColor = null;
            sSize = null;
        }

        public Symbol Clone()
        {
            return (Symbol)this.MemberwiseClone();
        }

        public void SetType(ref Object type)
        {
            this.sType = type;
        }
        public void SetColor(ref Object color)
        {
            this.sColor = color;
        }
        public void SetSize(ref Object size)
        {
            this.sSize = size;
        }
        public void SetDefault()
        {
            sType = null;
            sColor = null;
            sSize = null;
        }
        public Object GetColor()
        {
            return this.sColor;
        }
        public Object GetSize()
        {
            return this.sSize;
        }
        public Object GetTypes()
        {
            return this.sType;
        }

        private String byCol;
        public void SetGroupBy(String byColumn)
        {
            this.byCol = byColumn;
        }


        private StringBuilder cmnd = new StringBuilder();
        public virtual String GetCommand()
        {
            cmnd.Clear();
            cmnd.Append(" SYMBOL" + (this.byCol == null ? ";" : " " + this.byCol + ";") + Environment.NewLine);
            Type t;
            if (this.sType != null)
            {
                t = this.sType.GetType();
                if (t.IsArray)
                {                    
                    try
                    {
                        IEnumerable em = this.sType as IEnumerable;
                        cmnd.Append("  TYPE ");
                        foreach (object o in em)
                        {
                            cmnd.Append(o.ToString() + " ");
                        }
                        cmnd.Append(";" + Environment.NewLine);

                    }
                    catch 
                    {
                        throw new ArgumentException("Invalid Type value of Symbol");
                    }
                }
                else
                {
                    Console.WriteLine("It is not array, {0}", this.sType.ToString());
                    cmnd.Append("  TYPE " + this.sType.ToString() + ";" + Environment.NewLine);

                }

            }
            if (this.sColor != null)
            {
                t = this.sColor.GetType();
                Console.WriteLine(t.ToString());
                if (t.IsArray)
                {
                    Console.WriteLine("It is array");
                    try
                    {
                        IEnumerable em = this.sColor as IEnumerable;
                        StringBuilder tmpSb = new StringBuilder();
                        cmnd.Append("  COLOR ");
                        foreach (object o in em)
                        {
                            cmnd.Append(o.ToString() + " ");
                        }
                        cmnd.Append(";" + Environment.NewLine);
                    }
                    catch
                    {
                        throw new ArgumentException("Invalid Color value of Symbol");
                    }
                }
                else
                {
                    cmnd.Append("  COLOR " + this.sColor.ToString() + ";" + Environment.NewLine);
                }

            }
            if (this.sSize != null)
            {
                t = this.sSize.GetType();
                Console.WriteLine(t.ToString());
                if (t.IsArray)
                {
                    Console.WriteLine("It is array");
                    try
                    {
                        IEnumerable em = this.sColor as IEnumerable;
                        StringBuilder tmpSb = new StringBuilder();
                        cmnd.Append("  SIZE ");
                        foreach (object o in em)
                        {
                            cmnd.Append(o.ToString() + " ");
                        }
                        cmnd.Append(";" + Environment.NewLine);
                    }
                    catch
                    {
                        throw new ArgumentException("Invalid Size value of Symbol");
                    }
                }
                else
                {
                    cmnd.Append("  SIZE " + this.sSize.ToString() + ";" + Environment.NewLine);
                }
            }

            return cmnd.ToString();
        }

    }

}
