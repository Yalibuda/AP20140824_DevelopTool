using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace MtbGraph.GraphComponent
{
    
    public interface ICOMInterop_Annotation
    {
        String Title { set; get; }
        float TitleFontSize { set; get; }
        void AddFootnote(String footnote);
        //void AddFootnote(String footnote, int color, bool italic);
        void RemoveFootnoteAt(int i);
        void ClearFootnote();
        void SetDefault();

    }
}
