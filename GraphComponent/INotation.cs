using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace MtbGraph.GraphComponent
{
    public interface INotation
    {
       
        string Text { set; get; }
        float Size { set; get; }
        bool Italic { set; get; }
        bool Bold { set; get; }
        int Color { set; get; }
        INotation Clone();
    }
}
