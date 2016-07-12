using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component
{
    public interface IFootnote
    {
        void AddFootnote(dynamic footnote);
        void RemoveAll();
        int FontColor { get; set; }
        float FontSize { get; set; }        
    }
}
