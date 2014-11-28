using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    [ClassInterface(ClassInterfaceType.None)] //自動接口
    public class Line : ICOMInterop_MainLine
    {
        public Symbol Symbols { private set; get; }
        public Connectline Connectlines { private set; get; }
        public Line()
        {
            Symbols = new Symbol();
            Connectlines = new Connectline();
        }

        public Line Clone()
        {
            Line line = new Line();
            line.Symbols = this.Symbols.Clone();
            line.Connectlines = this.Connectlines.Clone();
            return line;
        }

        private StringBuilder cmnd = new StringBuilder();
        public virtual String GetCommand()
        {
            cmnd.Clear();
            cmnd.Append(Symbols.GetCommand());
            cmnd.Append(Connectlines.GetCommand());
            return cmnd.ToString();
        }

    }
}
