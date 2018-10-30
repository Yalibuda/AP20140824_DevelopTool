using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component
{
    public interface ILabelPosition
    {
        bool Visible { set; get; }
        float FontSize { set; get; }
        int FontColor { get; set; }
        int Model { set; get; }
        int[] RowId { set; get; }
        string Text { set; get; }
    }
}
