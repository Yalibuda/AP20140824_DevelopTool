using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.InteropServices;

namespace MtbGraph.GraphComponent
{
    public interface ICOMInterop_Line
    {
        //void SetType(ref Object type);

        /* SetColor
         * 20131204: 避免與Trend 繪製衝突先關閉
         * 20150129: 開啟功能，如果要使用就必須指定所有的symbol color
         */ 
        //void SetColor(ref Object color); 
        //void SetSize(ref Object size);
        void SetType(dynamic type);
        void SetColor(dynamic color);
        void SetSize(dynamic size);
        void SetDefault();
    }
}
