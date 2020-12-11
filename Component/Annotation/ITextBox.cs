using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Annotation
{
    public interface ITextBox
    {
        string[] Boxposition { get; set; }
        string Text { get; set; }
        int Unit { get; set; }
        void SetBoxposition(params object[] args);
        void SetCoordinate(params object[] args);
        string[] GetCoordinate();
        string GetCommand();
        // textbox position
    }
}
