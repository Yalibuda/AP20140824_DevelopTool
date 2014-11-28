using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MtbGraph.GraphComponent
{
    public interface IDataView
    {
        void SetType(ref Object type);
        Object GetTypes();
        void SetColor(ref Object color);
        Object GetColor();
        void SetSize(ref Object size);
        Object GetSize();
        void SetDefault();
        String GetCommand();
    }
}
