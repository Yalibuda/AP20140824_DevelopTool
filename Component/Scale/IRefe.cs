using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.Component.Scale
{
    public interface IRefe
    {
        void AddValue(double value);        
        //void RemoveAt(int i);
        void RemoveAll();
        float FontSize { set; get; }
        void SetType(dynamic var);
        void SetColor(dynamic var);
        void SetSize(dynamic var);
    }
}
