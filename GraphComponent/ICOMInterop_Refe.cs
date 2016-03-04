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
        void SetValue(ref Object value); //這裡設定 object 是要讓 Vb6 的陣列可以進來
        void SetColor(ref Object value);
        void SetType(ref Object value);
        bool HideLabel { set; get; }
        int FontSize { set; get; } //統一 label 字體大小
        int Size { set; get; } // 統一線的寬度
        void Clear();
    }
}
