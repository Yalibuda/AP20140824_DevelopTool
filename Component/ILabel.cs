using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component
{
    public interface ILabel
    {
        bool Visible { set; get; }
        string Text { set; get; }
        float FontSize { set; get; }
    }
}
