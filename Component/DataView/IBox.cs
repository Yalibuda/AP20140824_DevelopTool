using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.DataView
{
    public interface IBox
    {
        bool Visible { set; get; }
        void SetEType(dynamic var);
        void SetEColor(dynamic var);
        void SetESize(dynamic var);        
        void SetType(dynamic var);
        void SetColor(dynamic var);
        
    }
}
