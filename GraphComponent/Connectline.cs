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
        private Object cType;
        private Object cColor;
        private Object cSize;
        public Connectline()
        {
            cType = null;
            cColor = null;
            cSize = null;
        }

        public Connectline Clone()
        {
            return (Connectline)this.MemberwiseClone();
        }

        public void SetType(ref Object type)
        {
            this.cType = type;
        }
        public void SetColor(ref Object color)
        {
            this.cColor = color;
        }
        public void SetSize(ref Object size)
        {
            this.cSize = size;
        }
        public void SetDefault()
        {
            cType = null;
            cColor = null;
            cSize = null;
        }

        public object GetTypes()
        {
            throw new NotImplementedException();
        }

        public object GetColor()
        {
            throw new NotImplementedException();
        }

        public object GetSize()
        {
            throw new NotImplementedException();
        }

        private StringBuilder cmnd = new StringBuilder();
        public virtual String GetCommand()
        {
            cmnd.Clear();
            cmnd.AppendLine(" CONN" + (this.byCol == null ? ";" : " " + this.byCol + ";"));
            Type t;
            if (this.cType != null)
            {
                t = this.cType.GetType();
                if (t.IsArray)
                {
                    Console.WriteLine("It is array");
                    try
                    {
                        IEnumerable em = this.cType as IEnumerable;
                        StringBuilder tmpSb = new StringBuilder();
                        cmnd.Append("  TYPE ");
                        foreach (object o in em)
                        {
                            cmnd.Append(o.ToString() + " ");
                        }
                        cmnd.AppendLine(";");

                    }
                    catch (Exception e)
                    {
                        throw new Exception(e.Message);
                    }
                }
                else
                {
                    Console.WriteLine("It is not array, {0}", this.cType.ToString());
                    cmnd.Append("  TYPE " + this.cType.ToString() + ";" + Environment.NewLine);
                }

            }
            if (this.cColor != null)
            {
                t = this.cColor.GetType();
                Console.WriteLine(t.ToString());
                if (t.IsArray)
                {
                    Console.WriteLine("It is array");
                    try
                    {
                        IEnumerable em = this.cColor as IEnumerable;
                        StringBuilder tmpSb = new StringBuilder();
                        cmnd.Append("  COLOR ");
                        foreach (object o in em)
                        {
                            cmnd.Append(o.ToString() + " ");
                        }
                        cmnd.Append(";" + Environment.NewLine);
                    }
                    catch (Exception e)
                    {
                        throw new Exception(e.Message);
                    }
                }
                else
                {
                    cmnd.Append("  COLOR " + this.cColor.ToString() + ";" + Environment.NewLine);
                }

            }
            if (this.cSize != null)
            {
                t = this.cSize.GetType();
                Console.WriteLine(t.ToString());
                if (t.IsArray)
                {
                    Console.WriteLine("It is array");
                    try
                    {
                        IEnumerable em = this.cSize as IEnumerable;
                        StringBuilder tmpSb = new StringBuilder();
                        cmnd.Append("  SIZE ");
                        foreach (object o in em)
                        {
                            cmnd.Append(o.ToString() + " ");
                        }
                        cmnd.Append(";" + Environment.NewLine);
                    }
                    catch (Exception e)
                    {
                        throw new Exception(e.Message);
                    }
                }
                else
                {
                    cmnd.Append("  SIZE " + this.cSize.ToString() + ";" + Environment.NewLine);
                }
            }

            return cmnd.ToString();
        }


        private String byCol;
        public void SetGroupBy(string byColumn)
        {
            this.byCol = byColumn;
        }
    }
}
