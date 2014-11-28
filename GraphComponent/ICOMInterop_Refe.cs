using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace MtbGraph.GraphComponent
{
    public interface ICOMInterop_Refe
    {
        void SetValue(ref Object value);
        void SetColor(ref Object value);
        void SetType(ref Object value);
        void Clear();
    }
}
